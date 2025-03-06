import pandas as pd
import numpy as np
import os

class Get_Data():
    def _read_file(self, file_type: str, **kwargs) -> pd.DataFrame:
        """统一文件读取逻辑"""
        read_method = {
            'xlsx': pd.read_excel,
            'csv': pd.read_csv
        }.get(file_type.lower())

        if not read_method:
            raise ValueError(f"不支持的文件类型: {file_type}，当前支持 [xlsx/csv]")  # 优化错误提示信息

        return read_method(self.fileDataUrl, **kwargs)

    def getFileData(self, fileDataUrl: str) -> pd.DataFrame:
        """读取单文件数据
        优化点：
        1. 使用更安全的扩展名提取方式
        2. 增加文件存在性校验
        """
        if not os.path.exists(fileDataUrl):
            raise FileNotFoundError(f"文件路径不存在: {fileDataUrl}")
        self.fileDataUrl = fileDataUrl
        file_type = os.path.splitext(self.fileDataUrl)[-1][1:]  # 使用标准库方法获取扩展名
        try:
            fileData = self._read_file(file_type)
        except Exception as e:
            raise RuntimeError(f"读取文件失败: {self.fileDataUrl}") from e
        return fileData



