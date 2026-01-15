"""Excel 加载与管理模块 - 支持多表管理

v2.0: 支持双引擎模式 (ExcelDocument)
"""

import shutil
import uuid
from dataclasses import dataclass, field
from datetime import datetime
from pathlib import Path
from typing import Any, Dict, List, Optional, Union

import pandas as pd

from .config import get_config
from .excel_document import ExcelDocument


@dataclass
class TableInfo:
    """表的元信息"""
    id: str
    filename: str
    file_path: str  # 工作文件路径（副本）
    original_path: str  # 原始文件路径
    sheet_name: str
    total_rows: int
    total_columns: int
    loaded_at: datetime = field(default_factory=datetime.now)
    is_joined: bool = False  # 是否为连接表
    source_tables: List[str] = field(default_factory=list)  # 源表名称列表
    use_dual_engine: bool = False  # 是否使用双引擎模式
    is_copy: bool = False  # 是否为副本文件


class ExcelLoader:
    """Excel 文件加载器"""
    
    def __init__(self):
        self._df: Optional[pd.DataFrame] = None
        self._file_path: Optional[str] = None
        self._sheet_name: Optional[str] = None
        self._all_sheets: List[str] = []
    
    @property
    def is_loaded(self) -> bool:
        """是否已加载文件"""
        return self._df is not None
    
    @property
    def dataframe(self) -> pd.DataFrame:
        """获取 DataFrame"""
        if self._df is None:
            raise ValueError("未加载 Excel 文件")
        return self._df
    
    def load(self, file_path: str, sheet_name: Optional[str] = None) -> Dict[str, Any]:
        """加载 Excel 文件
        
        Args:
            file_path: Excel 文件路径
            sheet_name: 工作表名称，默认加载第一个
            
        Returns:
            文件结构信息
        """
        path = Path(file_path)
        if not path.exists():
            raise FileNotFoundError(f"文件不存在: {file_path}")
        
        if path.suffix.lower() not in ['.xlsx', '.xls', '.xlsm']:
            raise ValueError(f"不支持的文件格式: {path.suffix}")
        
        # 获取所有工作表名称
        xlsx = pd.ExcelFile(file_path)
        self._all_sheets = xlsx.sheet_names
        
        # 确定要加载的工作表
        if sheet_name is None:
            sheet_name = self._all_sheets[0]
        elif sheet_name not in self._all_sheets:
            raise ValueError(f"工作表 '{sheet_name}' 不存在，可用工作表: {self._all_sheets}")
        
        # 加载数据
        self._df = pd.read_excel(file_path, sheet_name=sheet_name)
        self._file_path = file_path
        self._sheet_name = sheet_name
        
        return self.get_structure()
    
    def get_structure(self) -> Dict[str, Any]:
        """获取 Excel 结构信息"""
        if self._df is None:
            raise ValueError("未加载 Excel 文件")
        
        config = get_config()
        
        # 列信息
        columns_info = []
        for col in self._df.columns:
            col_data = self._df[col]
            dtype = str(col_data.dtype)
            non_null = col_data.count()
            null_count = col_data.isna().sum()
            
            columns_info.append({
                "name": str(col),
                "dtype": dtype,
                "non_null_count": int(non_null),
                "null_count": int(null_count),
            })
        
        return {
            "file_path": self._file_path,
            "sheet_name": self._sheet_name,
            "all_sheets": self._all_sheets,
            "total_rows": len(self._df),
            "total_columns": len(self._df.columns),
            "columns": columns_info,
        }
    
    def get_preview(self, n_rows: Optional[int] = None) -> Dict[str, Any]:
        """获取数据预览
        
        Args:
            n_rows: 预览行数，默认使用配置值
            
        Returns:
            预览数据
        """
        if self._df is None:
            raise ValueError("未加载 Excel 文件")
        
        config = get_config()
        if n_rows is None:
            n_rows = config.excel.max_preview_rows
        
        preview_df = self._df.head(n_rows)
        
        return {
            "columns": list(self._df.columns),
            "data": preview_df.to_dict(orient="records"),
            "preview_rows": len(preview_df),
            "total_rows": len(self._df),
        }
    
    def switch_sheet(self, sheet_name: str) -> Dict[str, Any]:
        """切换到指定工作表

        Args:
            sheet_name: 目标工作表名称

        Returns:
            切换后的结构信息
        """
        if self._file_path is None:
            raise ValueError("未加载 Excel 文件")

        if sheet_name not in self._all_sheets:
            raise ValueError(f"工作表 '{sheet_name}' 不存在，可用工作表: {self._all_sheets}")

        if sheet_name == self._sheet_name:
            return self.get_structure()

        # 重新加载指定工作表
        self._df = pd.read_excel(self._file_path, sheet_name=sheet_name)
        self._sheet_name = sheet_name

        return self.get_structure()

    def get_all_sheets(self) -> List[str]:
        """获取所有工作表名称"""
        return self._all_sheets.copy()

    def get_summary(self) -> str:
        """获取 Excel 摘要信息（用于 Agent 上下文）"""
        if self._df is None:
            return "未加载 Excel 文件"
        
        structure = self.get_structure()
        preview = self.get_preview()
        
        lines = [
            f"📊 **已加载 Excel 文件**: {structure['file_path']}",
            f"📋 **当前工作表**: {structure['sheet_name']}",
            f"📑 **所有工作表**: {', '.join(structure['all_sheets'])}",
            f"📏 **数据规模**: {structure['total_rows']} 行 × {structure['total_columns']} 列",
            "",
            "**列信息**:",
        ]
        
        for col in structure['columns']:
            lines.append(f"  - `{col['name']}` ({col['dtype']}): {col['non_null_count']} 非空值")
        
        lines.append("")
        lines.append(f"**前 {preview['preview_rows']} 行数据预览**:")
        
        # 简单表格格式
        if preview['data']:
            headers = preview['columns']
            lines.append("| " + " | ".join(str(h) for h in headers) + " |")
            lines.append("| " + " | ".join("---" for _ in headers) + " |")
            for row in preview['data']:
                values = [str(row.get(h, ""))[:20] for h in headers]  # 截断长值
                lines.append("| " + " | ".join(values) + " |")
        
        return "\n".join(lines)


class MultiExcelLoader:
    """多表管理器 - 管理多个 Excel 文档实例

    v2.0: 支持双引擎模式，可选择使用 ExcelDocument（支持写入）或 ExcelLoader（只读）
    """

    def __init__(self, use_dual_engine: bool = True):
        """初始化多表管理器

        Args:
            use_dual_engine: 是否默认使用双引擎模式（支持写入）
        """
        self._tables: Dict[str, Union[ExcelLoader, ExcelDocument]] = {}
        self._table_infos: Dict[str, TableInfo] = {}
        self._active_table_id: Optional[str] = None
        self._default_dual_engine: bool = use_dual_engine
    
    @property
    def is_loaded(self) -> bool:
        """是否有任何表已加载"""
        return len(self._tables) > 0
    
    @property
    def active_table_id(self) -> Optional[str]:
        """获取当前活跃表ID"""
        return self._active_table_id
    
    def add_table(
        self,
        file_path: str,
        sheet_name: Optional[str] = None,
        use_dual_engine: Optional[bool] = None,
        create_copy: bool = False
    ) -> tuple[str, Dict[str, Any]]:
        """添加一张新表

        Args:
            file_path: Excel 文件路径
            sheet_name: 工作表名称
            use_dual_engine: 是否使用双引擎模式，默认使用初始化时的设置
            create_copy: 是否创建副本保护原始文件，默认 False（建议用户自行备份）

        Returns:
            (表ID, 结构信息)
        """
        original_path = Path(file_path)
        if not original_path.exists():
            raise FileNotFoundError(f"文件不存在: {file_path}")

        # 确定是否使用双引擎
        dual_engine = use_dual_engine if use_dual_engine is not None else self._default_dual_engine

        # 生成唯一ID
        table_id = str(uuid.uuid4())[:8]

        # 创建副本文件
        work_path = file_path
        is_copy = False

        if create_copy:
            # 创建副本保护原始文件
            copy_dir = original_path.parent / ".excel_copies"
            copy_dir.mkdir(exist_ok=True)

            # 副本文件名: 原文件名_copy_表ID.xlsx
            stem = original_path.stem
            suffix = original_path.suffix
            copy_filename = f"{stem}_copy_{table_id}{suffix}"
            copy_path = copy_dir / copy_filename

            # 复制文件
            shutil.copy2(file_path, copy_path)
            work_path = str(copy_path)
            is_copy = True
            print(f"[文件保护] 已创建副本: {copy_path}")

        # 创建加载器并加载数据
        if dual_engine:
            loader = ExcelDocument()
        else:
            loader = ExcelLoader()

        structure = loader.load(work_path, sheet_name)

        # 获取文件名
        filename = original_path.name

        # 存储表信息
        self._tables[table_id] = loader
        self._table_infos[table_id] = TableInfo(
            id=table_id,
            filename=filename,
            file_path=work_path,
            original_path=str(original_path.absolute()),
            sheet_name=structure["sheet_name"],
            total_rows=structure["total_rows"],
            total_columns=structure["total_columns"],
            use_dual_engine=dual_engine,
            is_copy=is_copy,
        )

        # 自动设为活跃表
        self._active_table_id = table_id

        return table_id, structure
    
    def remove_table(self, table_id: str) -> bool:
        """删除指定表

        Args:
            table_id: 表ID

        Returns:
            是否删除成功
        """
        if table_id not in self._tables:
            return False

        # 获取表信息，清理副本文件
        table_info = self._table_infos.get(table_id)
        if table_info and table_info.is_copy:
            try:
                copy_path = Path(table_info.file_path)
                if copy_path.exists():
                    copy_path.unlink()
                    print(f"[文件保护] 已删除副本: {copy_path}")
            except Exception as e:
                print(f"[文件保护] 删除副本失败: {e}")

        del self._tables[table_id]
        del self._table_infos[table_id]

        # 如果删除的是活跃表，切换到另一张表或设为None
        if self._active_table_id == table_id:
            if self._tables:
                self._active_table_id = next(iter(self._tables.keys()))
            else:
                self._active_table_id = None

        return True
    
    def get_table(self, table_id: str) -> Optional[Union[ExcelLoader, ExcelDocument]]:
        """获取指定表的加载器"""
        return self._tables.get(table_id)

    def get_document(self, table_id: str) -> Optional[ExcelDocument]:
        """获取指定表的 ExcelDocument（仅双引擎模式）"""
        table = self._tables.get(table_id)
        if isinstance(table, ExcelDocument):
            return table
        return None
    
    def get_table_info(self, table_id: str) -> Optional[TableInfo]:
        """获取指定表的元信息"""
        return self._table_infos.get(table_id)
    
    def get_active_loader(self) -> Optional[Union[ExcelLoader, ExcelDocument]]:
        """获取当前活跃表的加载器"""
        if self._active_table_id:
            return self._tables.get(self._active_table_id)
        return None

    def get_active_document(self) -> Optional[ExcelDocument]:
        """获取当前活跃表的 ExcelDocument（仅双引擎模式）"""
        if self._active_table_id:
            table = self._tables.get(self._active_table_id)
            if isinstance(table, ExcelDocument):
                return table
        return None

    def is_dual_engine(self, table_id: Optional[str] = None) -> bool:
        """检查指定表是否使用双引擎模式

        Args:
            table_id: 表ID，默认检查当前活跃表
        """
        table_id = table_id or self._active_table_id
        if table_id:
            info = self._table_infos.get(table_id)
            return info.use_dual_engine if info else False
        return False
    
    def get_active_table_info(self) -> Optional[TableInfo]:
        """获取当前活跃表的元信息"""
        if self._active_table_id:
            return self._table_infos.get(self._active_table_id)
        return None
    
    def set_active_table(self, table_id: str) -> bool:
        """设置当前活跃表
        
        Args:
            table_id: 表ID
            
        Returns:
            是否设置成功
        """
        if table_id not in self._tables:
            return False
        self._active_table_id = table_id
        return True
    
    def list_tables(self) -> List[Dict[str, Any]]:
        """获取所有表的信息列表"""
        result = []
        for table_id, info in self._table_infos.items():
            table = self._tables.get(table_id)
            is_dirty = False
            if isinstance(table, ExcelDocument):
                is_dirty = table.is_dirty

            result.append({
                "id": info.id,
                "filename": info.filename,
                "sheet_name": info.sheet_name,
                "total_rows": info.total_rows,
                "total_columns": info.total_columns,
                "loaded_at": info.loaded_at.isoformat(),
                "is_active": table_id == self._active_table_id,
                "is_joined": info.is_joined,
                "source_tables": info.source_tables,
                "use_dual_engine": info.use_dual_engine,
                "is_dirty": is_dirty,
            })
        return result
    
    def get_table_columns(self, table_id: str) -> List[str]:
        """获取指定表的列名列表"""
        loader = self.get_table(table_id)
        if loader and loader.is_loaded:
            return list(loader.dataframe.columns)
        return []
    
    def join_tables(
        self,
        table1_id: str,
        table2_id: str,
        keys1: List[str],
        keys2: List[str],
        join_type: str = "inner",
        new_name: str = "连接表"
    ) -> tuple[str, Dict[str, Any]]:
        """连接两张表（支持多字段连接）
        
        Args:
            table1_id: 表1 ID
            table2_id: 表2 ID
            keys1: 表1 连接字段列表
            keys2: 表2 连接字段列表
            join_type: 连接类型 (inner/left/right/outer)
            new_name: 新表名称
            
        Returns:
            (新表ID, 结构信息)
        """
        # 验证表存在
        loader1 = self.get_table(table1_id)
        loader2 = self.get_table(table2_id)
        if not loader1 or not loader2:
            raise ValueError("指定的表不存在")
        
        info1 = self.get_table_info(table1_id)
        info2 = self.get_table_info(table2_id)
        
        df1 = loader1.dataframe
        df2 = loader2.dataframe
        
        # 验证字段数量一致
        if len(keys1) != len(keys2):
            raise ValueError("两表的连接字段数量必须一致")
        
        if len(keys1) == 0:
            raise ValueError("至少需要指定一个连接字段")
        
        # 验证字段存在
        for key in keys1:
            if key not in df1.columns:
                raise ValueError(f"表1中不存在字段: {key}")
        for key in keys2:
            if key not in df2.columns:
                raise ValueError(f"表2中不存在字段: {key}")
        
        # 验证连接类型
        valid_join_types = ["inner", "left", "right", "outer"]
        if join_type not in valid_join_types:
            raise ValueError(f"不支持的连接类型: {join_type}，可选: {valid_join_types}")
        
        # 执行连接
        merged_df = pd.merge(
            df1, df2,
            left_on=keys1,
            right_on=keys2,
            how=join_type,
            suffixes=('_表1', '_表2')
        )
        
        # 创建新的加载器
        new_loader = ExcelLoader()
        new_loader._df = merged_df
        new_loader._file_path = f"[连接表] {new_name}"
        new_loader._sheet_name = "merged"
        new_loader._all_sheets = ["merged"]
        
        # 生成唯一ID
        table_id = str(uuid.uuid4())[:8]
        
        # 存储表信息
        self._tables[table_id] = new_loader
        self._table_infos[table_id] = TableInfo(
            id=table_id,
            filename=f"🔗 {new_name}",
            file_path=f"[连接表] {new_name}",
            sheet_name="merged",
            total_rows=len(merged_df),
            total_columns=len(merged_df.columns),
            is_joined=True,
            source_tables=[info1.filename, info2.filename],
        )
        
        # 自动设为活跃表
        self._active_table_id = table_id
        
        return table_id, new_loader.get_structure()
    
    def get_active_summary(self) -> str:
        """获取当前活跃表的摘要"""
        loader = self.get_active_loader()
        if loader:
            return loader.get_summary()
        return "未加载 Excel 文件"
    
    def get_summary(self) -> str:
        """获取当前活跃表的摘要（兼容旧接口）"""
        return self.get_active_summary()
    
    @property
    def dataframe(self) -> pd.DataFrame:
        """获取当前活跃表的 DataFrame（兼容旧接口）"""
        loader = self.get_active_loader()
        if loader:
            return loader.dataframe
        raise ValueError("未加载 Excel 文件")

    # ==================== 写入操作（仅双引擎模式） ====================

    def save_table(
        self,
        table_id: Optional[str] = None,
        file_path: Optional[str] = None
    ) -> Dict[str, Any]:
        """保存指定表（保存到副本文件）

        Args:
            table_id: 表ID，默认当前活跃表
            file_path: 保存路径，默认保存到副本文件

        Returns:
            操作结果
        """
        table_id = table_id or self._active_table_id
        if not table_id:
            raise ValueError("未指定表ID且无活跃表")

        doc = self.get_document(table_id)
        if not doc:
            raise ValueError("该表不支持保存操作（非双引擎模式）")

        save_path = doc.save(file_path)
        return {
            "success": True,
            "table_id": table_id,
            "file_path": save_path
        }

    def save_to_original(
        self,
        table_id: Optional[str] = None
    ) -> Dict[str, Any]:
        """将修改保存回原始文件（覆盖原文件）

        警告：此操作会覆盖原始文件，请确保已备份重要数据。

        Args:
            table_id: 表ID，默认当前活跃表

        Returns:
            操作结果
        """
        table_id = table_id or self._active_table_id
        if not table_id:
            raise ValueError("未指定表ID且无活跃表")

        table_info = self._table_infos.get(table_id)
        if not table_info:
            raise ValueError(f"表不存在: {table_id}")

        doc = self.get_document(table_id)
        if not doc:
            raise ValueError("该表不支持保存操作（非双引擎模式）")

        # 保存到原始文件路径
        original_path = table_info.original_path
        save_path = doc.save(original_path)

        return {
            "success": True,
            "table_id": table_id,
            "original_path": original_path,
            "file_path": save_path,
            "message": f"已将修改保存到原始文件: {original_path}"
        }

    def export_to(
        self,
        export_path: str,
        table_id: Optional[str] = None
    ) -> Dict[str, Any]:
        """导出到新文件

        Args:
            export_path: 导出路径
            table_id: 表ID，默认当前活跃表

        Returns:
            操作结果
        """
        table_id = table_id or self._active_table_id
        if not table_id:
            raise ValueError("未指定表ID且无活跃表")

        doc = self.get_document(table_id)
        if not doc:
            raise ValueError("该表不支持导出操作（非双引擎模式）")

        save_path = doc.save(export_path)
        return {
            "success": True,
            "table_id": table_id,
            "export_path": save_path,
            "message": f"已导出到: {export_path}"
        }

    def write_cell(
        self,
        cell: str,
        value: Any,
        table_id: Optional[str] = None,
        sheet: Optional[str] = None
    ) -> Dict[str, Any]:
        """写入单元格

        Args:
            cell: 单元格地址，如 "A1"
            value: 写入的值
            table_id: 表ID，默认当前活跃表
            sheet: 工作表名称，默认当前表

        Returns:
            操作结果
        """
        table_id = table_id or self._active_table_id
        if not table_id:
            raise ValueError("未指定表ID且无活跃表")

        doc = self.get_document(table_id)
        if not doc:
            raise ValueError("该表不支持写入操作（非双引擎模式）")

        return doc.write_cell(cell, value, sheet)

    def write_range(
        self,
        start_cell: str,
        data: List[List[Any]],
        table_id: Optional[str] = None,
        sheet: Optional[str] = None
    ) -> Dict[str, Any]:
        """批量写入数据

        Args:
            start_cell: 起始单元格，如 "A1"
            data: 二维数据数组
            table_id: 表ID，默认当前活跃表
            sheet: 工作表名称，默认当前表

        Returns:
            操作结果
        """
        table_id = table_id or self._active_table_id
        if not table_id:
            raise ValueError("未指定表ID且无活跃表")

        doc = self.get_document(table_id)
        if not doc:
            raise ValueError("该表不支持写入操作（非双引擎模式）")

        return doc.write_range(start_cell, data, sheet)

    def write_formula(
        self,
        cell: str,
        formula: str,
        table_id: Optional[str] = None,
        sheet: Optional[str] = None
    ) -> Dict[str, Any]:
        """写入公式

        Args:
            cell: 单元格地址
            formula: 公式字符串，如 "SUM(A1:A10)" 或 "=SUM(A1:A10)"
            table_id: 表ID，默认当前活跃表
            sheet: 工作表名称，默认当前表

        Returns:
            操作结果
        """
        table_id = table_id or self._active_table_id
        if not table_id:
            raise ValueError("未指定表ID且无活跃表")

        doc = self.get_document(table_id)
        if not doc:
            raise ValueError("该表不支持写入操作（非双引擎模式）")

        return doc.write_formula(cell, formula, sheet)

    def read_formula(
        self,
        cell: str,
        table_id: Optional[str] = None,
        sheet: Optional[str] = None
    ) -> Optional[str]:
        """读取单元格公式

        Args:
            cell: 单元格地址
            table_id: 表ID，默认当前活跃表
            sheet: 工作表名称，默认当前表

        Returns:
            公式字符串，如果不是公式则返回 None
        """
        table_id = table_id or self._active_table_id
        if not table_id:
            raise ValueError("未指定表ID且无活跃表")

        doc = self.get_document(table_id)
        if not doc:
            raise ValueError("该表不支持公式读取（非双引擎模式）")

        return doc.read_formula(cell, sheet)

    def insert_rows(
        self,
        row: int,
        count: int = 1,
        table_id: Optional[str] = None,
        sheet: Optional[str] = None
    ) -> Dict[str, Any]:
        """插入行"""
        table_id = table_id or self._active_table_id
        if not table_id:
            raise ValueError("未指定表ID且无活跃表")

        doc = self.get_document(table_id)
        if not doc:
            raise ValueError("该表不支持此操作（非双引擎模式）")

        return doc.insert_rows(row, count, sheet)

    def delete_rows(
        self,
        start_row: int,
        end_row: Optional[int] = None,
        table_id: Optional[str] = None,
        sheet: Optional[str] = None
    ) -> Dict[str, Any]:
        """删除行"""
        table_id = table_id or self._active_table_id
        if not table_id:
            raise ValueError("未指定表ID且无活跃表")

        doc = self.get_document(table_id)
        if not doc:
            raise ValueError("该表不支持此操作（非双引擎模式）")

        return doc.delete_rows(start_row, end_row, sheet)

    def get_change_log(self, table_id: Optional[str] = None) -> List[Dict[str, Any]]:
        """获取变更日志

        Args:
            table_id: 表ID，默认当前活跃表

        Returns:
            变更记录列表
        """
        table_id = table_id or self._active_table_id
        if not table_id:
            return []

        doc = self.get_document(table_id)
        if not doc:
            return []

        return [
            {
                "type": change.change_type.value,
                "sheet": change.sheet_name,
                "location": change.location,
                "old_value": change.old_value,
                "new_value": change.new_value,
                "timestamp": change.timestamp.isoformat()
            }
            for change in doc.change_log
        ]


# 全局实例 - 使用多表管理器
_loader: Optional[MultiExcelLoader] = None


def get_loader(use_dual_engine: bool = True) -> MultiExcelLoader:
    """获取全局 MultiExcelLoader 实例

    Args:
        use_dual_engine: 是否使用双引擎模式（仅首次创建时生效）
    """
    global _loader
    if _loader is None:
        _loader = MultiExcelLoader(use_dual_engine=use_dual_engine)
    return _loader


def reset_loader(use_dual_engine: bool = True) -> None:
    """重置全局 MultiExcelLoader 实例

    Args:
        use_dual_engine: 新实例是否使用双引擎模式
    """
    global _loader
    _loader = MultiExcelLoader(use_dual_engine=use_dual_engine)
