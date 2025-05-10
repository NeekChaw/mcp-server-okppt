from pptx.util import Inches, Cm, Pt
from typing import Union, Optional

class UnitConverter:
    """单位转换工具类，用于处理PPT中的各种单位转换"""
    # EMU (English Metric Unit) 转换常量
    EMU_PER_INCH = 914400
    EMU_PER_CM = 360000
    EMU_PER_PT = 12700
    EMU_PER_MM = 36000
    EMU_PER_PX = 9525
    @staticmethod
    def to_emu(value: Union[int, float, str, Pt, Inches, Cm], total: Optional[float] = None) -> str:
        """
        将各种单位转换为EMU（English Metric Unit）
        
        Args:
            value: 要转换的值，可以是:
                - 数字（假定为EMU）
                - 字符串（支持以下格式：
                    - "50%" - 百分比（需要提供total参数）
                    - "12pt" - 磅值
                    - "1inch" - 英寸
                    - "2.54cm" - 厘米
                    - "25.4mm" - 毫米
                    - "1000px" - 像素
                    - "1000" - 纯数字假定为EMU）
                - Pt对象（磅值）
                - Inches对象（英寸）
                - Cm对象（厘米）
            total: 计算百分比时的总量（EMU单位）
        
        Returns:
            str: EMU值的字符串表示
            
        Note:
            EMU (English Metric Units) 转换关系:
            - 1 pt = 12700 EMU
            - 1 inch = 914400 EMU
            - 1 cm = 360000 EMU
            - 1 mm = 36000 EMU
            - 1 px = 9525 EMU
        """
        if isinstance(value, str):
            # 处理百分比
            if value.endswith('%') and total is not None:
                percent = float(value.strip('%')) / 100
                return str(int(percent * total))
            
            # 处理带单位的值
            import re
            match = re.match(r"(-?[\d.]+)([a-zA-Z]+)?", value)
            if not match:
                raise ValueError(f"无效的输入格式: {value}")
            
            number = float(match.group(1))
            unit = match.group(2)
            
            # 如果没有单位，假定为EMU
            if unit is None:
                return str(int(number))
            
            # 有单位的情况下再转小写
            unit = unit.lower()
            
            # 直接进行单位换算
            if unit == 'pt':
                return str(int(number * UnitConverter.EMU_PER_PT))
            elif unit == 'inch' or unit == 'in':
                return str(int(number * UnitConverter.EMU_PER_INCH))
            elif unit == 'cm':
                return str(int(number * UnitConverter.EMU_PER_CM))
            elif unit == 'mm':
                return str(int(number * UnitConverter.EMU_PER_MM))
            elif unit == 'px':
                return str(int(number * UnitConverter.EMU_PER_PX))
            else:
                raise ValueError(f"不支持的单位: {unit}")
            
        # 处理对象类型
        elif isinstance(value, (Pt, Inches, Cm)):
            # 获取对象的内部值，避免双重转换
            raw_value = float(value)
            if isinstance(value, Pt):
                return str(int(raw_value * 12700))
            elif isinstance(value, Inches):
                return str(int(raw_value * 914400))
            elif isinstance(value, Cm):  # Cm
                return str(int(raw_value * 360000))
            else:
                raise ValueError(f"不支持的单位: {type(value)}")

        # 处理数字类型（假定为EMU）
        else:
            return str(int(value))
        
    @staticmethod
    def to_inches(value: Union[int, float, str], total: Optional[float] = None) -> Inches:
        """
        将各种输入转换为英寸单位
        
        Args:
            value: 输入值，支持:
                - 数字（假定为英寸）
                - 字符串格式：
                    - "50%" - 百分比（需要提供total参数）
                    - "12pt" - 磅值
                    - "1inch" - 英寸
                    - "2.54cm" - 厘米
                    - "25.4mm" - 毫米
                    - "1000px" - 像素
                total: 计算百分比时的总量（英寸）
        
        Returns:
            Inches: 转换后的英寸值
        """
        if isinstance(value, str):
            # 处理百分比
            if value.endswith('%') and total is not None:
                percent = float(value.strip('%')) / 100
                return Inches(percent * total)
            
            # 处理带单位的值
            import re
            match = re.match(r"(-?[\d.]+)([a-zA-Z]+)?", value)
            if not match:
                raise ValueError(f"无效的输入格式: {value}")
            
            number = float(match.group(1))
            unit = match.group(2)
            
            # 如果没有单位，假定为英寸
            if unit is None:
                return Inches(number)
            
            # 有单位的情况下再转小写
            unit = unit.lower()
            
            # 直接进行单位换算
            if unit == 'pt':
                return Inches(number / 72.0)  # 1 inch = 72 pt
            elif unit == 'inch' or unit == 'in':
                return Inches(number)
            elif unit == 'cm':
                return Inches(number / 2.54)  # 1 inch = 2.54 cm
            elif unit == 'mm':
                return Inches(number / 25.4)  # 1 inch = 25.4 mm
            elif unit == 'px':
                return Inches(number / 9525)  # 1 inch = 9525 px
            else:
                raise ValueError(f"不支持的单位: {unit}")
            
        # 处理数字类型（假定为英寸）
        return Inches(float(value))

    @staticmethod
    def to_cm(value: Union[int, float, str], total: Optional[float] = None) -> Cm:
        """
        将各种输入转换为厘米单位
        
        Args:
            value: 输入值，支持:
                - 数字（假定为厘米）
                - 字符串格式：
                    - "50%" - 百分比（需要提供total参数）
                    - "12pt" - 磅值
                    - "1inch" - 英寸
                    - "2.54cm" - 厘米
                    - "25.4mm" - 毫米
                    - "1000px" - 像素
                total: 计算百分比时的总量（厘米）
        
        Returns:
            Cm: 转换后的厘米值
        """
        if isinstance(value, str):
            # 处理百分比
            if value.endswith('%') and total is not None:
                percent = float(value.strip('%')) / 100
                return Cm(percent * total)
            
            # 处理带单位的值
            import re
            match = re.match(r"(-?[\d.]+)([a-zA-Z]+)?", value)
            if not match:
                raise ValueError(f"无效的输入格式: {value}")
            
            number = float(match.group(1))
            unit = match.group(2)
            
            # 如果没有单位，假定为厘米
            if unit is None:
                return Cm(number)
            
            # 有单位的情况下再转小写
            unit = unit.lower()
            
            # 直接进行单位换算
            if unit == 'pt':
                return Cm(number * 2.54 / 72.0)  # 先转英寸再转厘米
            elif unit == 'inch' or unit == 'in':
                return Cm(number * 2.54)  # 1 inch = 2.54 cm
            elif unit == 'cm':
                return Cm(number)
            elif unit == 'mm':
                return Cm(number / 10)  # 1 cm = 10 mm
            elif unit == 'px':
                return Cm(number / 9525)  # 1 cm = 9525 px
            else:
                raise ValueError(f"不支持的单位: {unit}")
            
        # 处理数字类型（假定为厘米）
        return Cm(float(value))

    @staticmethod
    def to_points(value: Union[int, float, str], total: Optional[float] = None) -> Pt:
        """
        将各种输入转换为磅值单位
        
        Args:
            value: 输入值，支持:
                - 数字（假定为磅值）
                - 字符串格式：
                    - "50%" - 百分比（需要提供total参数）
                    - "12pt" - 磅值
                    - "1inch" - 英寸
                    - "2.54cm" - 厘米
                    - "25.4mm" - 毫米
                    - "1000px" - 像素
                total: 计算百分比时的总量（磅值）
        
        Returns:
            Pt: 转换后的磅值
        """
        if isinstance(value, str):
            # 处理百分比
            if value.endswith('%') and total is not None:
                percent = float(value.strip('%')) / 100
                return Pt(percent * total)
            
            # 处理带单位的值
            import re
            match = re.match(r"(-?[\d.]+)([a-zA-Z]+)?", value)
            if not match:
                raise ValueError(f"无效的输入格式: {value}")
            
            number = float(match.group(1))
            unit = match.group(2)
            
            # 如果没有单位，假定为磅值
            if unit is None:
                return Pt(number)
            
            # 有单位的情况下再转小写
            unit = unit.lower()
            
            # 直接进行单位换算
            if unit == 'pt':
                return Pt(number)
            elif unit == 'inch' or unit == 'in':
                return Pt(number * 72.0)  # 1 inch = 72 pt
            elif unit == 'cm':
                return Pt(number * 72.0 / 2.54)  # 先转英寸再转磅值
            elif unit == 'mm':
                return Pt(number * 72.0 / 25.4)  # 先转英寸再转磅值
            elif unit == 'px':
                return Pt(number * 72.0 / 9525)  # 先转英寸再转磅值
            else:
                raise ValueError(f"不支持的单位: {unit}")
            
        # 处理数字类型（假定为磅值）
        return Pt(float(value))

    @staticmethod
    def inches_to_cm(inches: float) -> float:
        """英寸转换为厘米"""
        return inches * 2.54

    @staticmethod
    def cm_to_inches(cm: float) -> float:
        """厘米转换为英寸"""
        return cm / 2.54

    @staticmethod
    def points_to_inches(points: float) -> float:
        """磅值转换为英寸"""
        return points / 72.0

    @staticmethod
    def inches_to_points(inches: float) -> float:
        """英寸转换为磅值"""
        return inches * 72.0
    
    @staticmethod
    def px_to_emu(px: float) -> float:
        """像素转换为EMU"""
        return px * 9525
    
    @staticmethod
    def emu_to_px(emu: float) -> float:
        """EMU转换为像素"""
        return emu / 9525
