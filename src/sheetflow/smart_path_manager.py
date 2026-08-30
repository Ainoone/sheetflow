"""时间敏感目录管理模块。

管理基于时间戳的目录创建和子项数量阈值检查。

主要特性：
    - 子项计数缓存：避免重复的文件系统扫描
    - 三分支决策：根据基目录位置和子项数量决定子目录创建位置
    - 桌面解析：Windows 读取用户重定向配置，其他平台使用现有 ~/Desktop
    - 原子目录创建：使用 exist_ok=False 检测并拒绝 TOCTOU 竞态冲突
    - 冲突处理：时间戳冲突时使用三级降级命名

典型用例：
    >>> from pathlib import Path
    >>> manager = SmartPathManager("/path/to/dir")
    >>> new_dir = manager.mkdir_if_needed()
    >>> if new_dir:
    ...     print(f"Created: {new_dir}")
"""

from datetime import datetime
import logging
import os
from pathlib import Path
import sys

# 依赖无项目内导入的叶子模块，避免路径层反向依赖 common。
from .user_errors import ERR_INVALID_PATH, format_user_error

logger = logging.getLogger(__name__)


class SmartPathManager:
    """管理时间敏感目录创建和子项数量阈值判断。

    三分支决策：基目录位置 + 子项数量 → 决定子目录创建位置。

    Attributes:
        DEFAULT_THRESHOLD: 默认子项数量阈值（类常量）
        MAX_SEQUENCE_ATTEMPTS: 序号后缀最大尝试次数（类常量）

    Args:
        path: 目标目录路径，支持 str 或 Path 对象

    Raises:
        ValueError: 当路径无效或不是目录时抛出

    Examples:
        >>> manager = SmartPathManager("/path/to/dir")
        >>> manager.count_files()
        2
        >>> # Case B: 子项数未达阈值 → 在基目录下创建子目录
        >>> manager.mkdir_if_needed(threshold=5)
        PosixPath('/path/to/dir/26_05_09_14')
        >>> # Case C: 子项数达到阈值 → 桌面可用时跳转桌面
        >>> manager.mkdir_if_needed(threshold=2)
        PosixPath('/home/user/Desktop/26_05_09_14')

    Note:
        - 子项计数结果会被缓存，使用 refresh=True 强制刷新
        - 子项计数忽略 Excel 临时文件（~$ 开头），与用户视角一致
        - 目录创建使用原子操作检测竞态冲突
        - 系统桌面无法解析或不存在时留在原目录，不创建猜测的 Desktop 目录
    """

    DEFAULT_THRESHOLD = 5
    MAX_SEQUENCE_ATTEMPTS = 100

    def __init__(self, path: str | Path) -> None:
        """初始化路径管理器。

        Args:
            path: 目标目录路径

        Raises:
            ValueError: 路径无效或不是目录
        """
        target = Path(path)
        if not target.exists() or not target.is_dir():
            raise ValueError(
                format_user_error(
                    ERR_INVALID_PATH,
                    f"Invalid directory path: {path}",
                    f"路径无效或不是目录：{path}",
                )
            )
        self._dir = target
        self._cached_count: int | None = None

    def count_files(self, *, refresh: bool = False) -> int:
        """统计目标目录直接包含的子项数量（文件 + 目录，不递归）。

        结果会被缓存直到传入 refresh=True 强制刷新。
        忽略 Excel 临时文件（~$ 开头），计数与用户视角一致。

        Args:
            refresh: 是否强制刷新缓存，默认 False

        Returns:
            目录中的子项数量（文件 + 目录）

        Examples:
            >>> manager = SmartPathManager("/path/to/dir")
            >>> manager.count_files()
            5
            >>> manager.count_files()  # 使用缓存
            5
            >>> manager.count_files(refresh=True)  # 强制刷新
            6

        Note:
            - 只统计直接子项，不递归子目录
        """
        if refresh or self._cached_count is None:
            self._cached_count = sum(
                1 for p in self._dir.iterdir()
                if not p.name.startswith('~$')
            )
        return self._cached_count

    @staticmethod
    def _desktop_path() -> Path | None:
        """返回系统实际桌面目录；无法可靠解析或目录不存在时返回 None。

        Windows 的桌面可能被 OneDrive 或企业策略重定向，因此读取当前用户的
        ``User Shell Folders\\Desktop`` 配置。macOS 等平台使用现有的
        ``~/Desktop``；不会为了满足约定而创建一个猜测路径。
        """
        if sys.platform == "win32":
            try:
                import winreg

                key_path = (
                    r"Software\Microsoft\Windows\CurrentVersion\Explorer"
                    r"\User Shell Folders"
                )
                with winreg.OpenKey(winreg.HKEY_CURRENT_USER, key_path) as key:
                    raw_path, _ = winreg.QueryValueEx(key, "Desktop")
                desktop = Path(os.path.expandvars(raw_path)).expanduser()
            except (ImportError, OSError, TypeError) as exc:
                logger.warning(f"无法读取 Windows 桌面路径: {exc}")
                return None
        else:
            desktop = Path.home() / "Desktop"

        if desktop.is_dir():
            return desktop

        logger.warning(f"系统桌面目录不存在或不是目录: {desktop}")
        return None

    def _is_desktop(self) -> bool:
        """判断当前管理的目录是否为用户桌面。"""
        desktop = self._desktop_path()
        if desktop is None:
            return False
        try:
            return self._dir.resolve() == desktop.resolve()
        except OSError as e:
            logger.warning(f"无法解析路径以判断是否为桌面: {e}")
            return False

    def _resolve_target_under(self, base: Path, now: datetime) -> Path | None:
        """计算基于时间戳的目标路径，支持冲突时的序号降级。

        命名策略：
            1. 优先使用小时级别:YY_MM_DD_HH (如 26_05_02_14)
            2. 如果小时目录已存在，使用分钟级别:YY_MM_DD_HHMM (如 26_05_02_1430)
            3. 如果分钟目录也存在，添加序号后缀:YY_MM_DD_HHMM_01, _02, ...

        Args:
            base: 基础目录路径
            now: 时间戳对象

        Returns:
            可用的目标路径，如果所有尝试都失败则返回 None

        Examples:
            >>> from datetime import datetime
            >>> manager = SmartPathManager("/path/to/dir")
            >>> now = datetime(2026, 5, 2, 14, 30, 45)
            >>> # 假设 26_05_02_14/ 不存在
            >>> manager._resolve_target_under(Path("/base"), now)
            PosixPath('/base/26_05_02_14')
            >>> # 假设 26_05_02_14/ 已存在
            >>> manager._resolve_target_under(Path("/base"), now)
            PosixPath('/base/26_05_02_1430')
            >>> # 假设 26_05_02_1430/ 也存在
            >>> manager._resolve_target_under(Path("/base"), now)
            PosixPath('/base/26_05_02_1430_01')

        Note:
            - 序号后缀最多尝试 MAX_SEQUENCE_ATTEMPTS 次
            - 所有候选路径都存在时返回 None
        """
        hour_name = now.strftime("%y_%m_%d_%H")
        minute_name = now.strftime("%y_%m_%d_%H%M")

        hour_dir = base / hour_name

        if not hour_dir.exists():
            return hour_dir

        minute_dir = base / minute_name
        if not minute_dir.exists():
            return minute_dir

        for seq in range(1, self.MAX_SEQUENCE_ATTEMPTS + 1):
            candidate = base / f"{minute_name}_{seq:02d}"
            if not candidate.exists():
                return candidate

        logger.error(
            f"无法在 {base} 下生成唯一目录名，已尝试 {self.MAX_SEQUENCE_ATTEMPTS} 次序号后缀"
        )
        return None

    def mkdir_if_needed(
        self,
        now: datetime | None = None,
        threshold: int = DEFAULT_THRESHOLD,
    ) -> Path | None:
        """按需创建目录，返回创建的路径或 None。

        Case A - 基目录是桌面：始终在桌面创建时间戳子目录。
        Case B - 基目录不是桌面，子项数 < threshold:在基目录下创建时间戳子目录。
        Case C - 基目录不是桌面，子项数 ≥ threshold:优先跳转系统桌面；桌面无法
        可靠解析时回退到基目录。

        Args:
            now: 时间戳对象，默认为 None(使用当前时间）
            threshold: 子项数量阈值，默认使用 DEFAULT_THRESHOLD (5)

        Returns:
            成功创建的目录路径；若所有候选目录名都已存在，或原子创建时恰好
            遇到同名目录，则返回 None

        Raises:
            OSError: 权限错误、磁盘错误等文件系统异常会继续向外抛出。

        Note:
            - 子项计数忽略 Excel 临时文件（~$ 开头），与用户视角一致
            - 使用 exist_ok=False 原子检测 TOCTOU 竞态；同名目标抢先出现时返回 None
            - Case C 只使用已存在的系统桌面；无法解析时继续在基目录创建
        """
        if now is None:
            now = datetime.now()

        if self._is_desktop():
            base = self._dir
            parents = False
        elif self.count_files(refresh=True) < threshold:
            base = self._dir
            parents = False
        else:
            # 达到阈值时优先桌面；桌面不可用则回退原目录。
            desktop = self._desktop_path()
            base = desktop if desktop is not None else self._dir
            parents = False

        target = self._resolve_target_under(base, now)
        if target is None:
            return None

        try:
            target.mkdir(parents=parents, exist_ok=False)
            return target
        except FileExistsError:
            return None
