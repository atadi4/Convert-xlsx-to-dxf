# -*- coding: gbk -*-   
"""
主程序：Excel坐标转DXF工具
说明：
可以直接运行，或者使用pyinstaller打包成exe文件。
如果要打包成exe文件，使用以下命令：pyinstaller --onefile --windowed --icon=doro.ico --add-data "doro.ico;." --add-data "My.jpg;." -n 铁匠一键转dxf工具 Mydxf.py
可以直接拿铁匠生成的else文件丢进去！！！
缺少的库自己看报错安装
功能：
1. 读取Excel文件中的坐标点（A列为X，B列为Y）
2. 在窗口中预览用直线连接的图形
3. 支持窗口自适应缩放、背景水印、右下角图片
4. 可导出为DXF文件
5. 支持自定义窗口和任务栏图标
"""

import sys
import os
from PyQt5.QtWidgets import QApplication, QMainWindow, QFileDialog, QPushButton, QWidget, QVBoxLayout
from PyQt5.QtGui import QPainter, QPen, QPixmap, QFont, QIcon
from PyQt5.QtCore import Qt
import openpyxl
import ezdxf

class GraphWidget(QWidget):
    """
    图形预览区，负责绘制背景水印、右下角图片和坐标连线
    """
    def __init__(self, parent=None):
        super().__init__(parent)
        # 加载右下角图片（兼容打包和源码运行）
        if getattr(sys, 'frozen', False):
            #右下角的图片要丢在目录下，名字为"My.jpg"
            img_path = os.path.join(sys._MEIPASS, "My.jpg")
        else:
            img_path = "My.jpg"
        self.icon = QPixmap(img_path)
        self.points = []

    def set_points(self, points):
        """
        设置要绘制的点，并刷新界面
        """
        self.points = points
        self.update()

    def paintEvent(self, event):
        """
        绘制事件：背景水印、右下角图片、连线
        """
        painter = QPainter(self)
        painter.fillRect(self.rect(), Qt.white)

        
        painter.setPen(Qt.gray)
        font = QFont("微软雅黑", 32)
        font.setItalic(True)
        painter.setFont(font)
        painter.setOpacity(0.15)  # 水印透明度
        text = "小蜜蜂实验室"
        painter.save()
        angle = -30  # 水印倾斜角度
        painter.rotate(angle)
        w = self.width()
        h = self.height()
        metrics = painter.fontMetrics()
        text_w = metrics.width(text)
        text_h = metrics.height()
        # 斜着铺满整个窗口
        x_start = -w
        x_end = w * 2
        y_start = 0
        y_end = h * 2
        step_x = text_w + 100
        step_y = text_h + 80
        for x in range(x_start, x_end, step_x):
            for y in range(y_start, y_end, step_y):
                painter.drawText(x, y, text)
        painter.restore()
        painter.setOpacity(1.0)  # 恢复不透明

        # 3. 绘制右下角图片（自动缩放）
        if not self.icon.isNull():
            icon_w = int(self.width() * 0.2)
            icon_h = int(self.icon.height() * icon_w / self.icon.width())
            icon_pix = self.icon.scaled(icon_w, icon_h, Qt.KeepAspectRatio, Qt.SmoothTransformation)
            x = self.width() - icon_w - 10
            y = self.height() - icon_h - 10
            painter.drawPixmap(x, y, icon_pix)

        # 4. 绘制坐标连线
        if not self.points:
            return
        pen = QPen(Qt.black, 2)
        painter.setPen(pen)
        for i in range(len(self.points) - 1):
            x1, y1 = self.points[i]
            x2, y2 = self.points[i + 1]
            painter.drawLine(x1, y1, x2, y2)

        # 5. 在图形中间绘制法兰盘（以(0,0)为圆心，直径12）
        import math
        cx, cy = self.width() // 2, self.height() // 2
        scale = min(self.width(), self.height()) / 80 

        # 中心大圆（直径12，圆心(0,0)）
        painter.setPen(QPen(Qt.blue, 2))
        painter.drawEllipse(
            int(cx - 6 * scale), int(cy - 6 * scale),
            int(12 * scale), int(12 * scale)
        )

        # 四个小孔（直径3，圆心分别为(10,0),(-10,0),(0,10),(0,-10)）
        painter.setPen(QPen(Qt.red, 2))
        for dx, dy in [(10, 0), (-10, 0), (0, 10), (0, -10)]:
            x = cx + dx * scale
            y = cy - dy * scale  # y轴向下为正，需取反
            painter.drawEllipse(
                int(x - 1.5 * scale), int(y - 1.5 * scale),
                int(3 * scale), int(3 * scale)
            )

    def resizeEvent(self, event):
        """
        窗口大小变化时自动缩放并刷新图形
        """
        main_window = self.parent().parent() if self.parent() else None
        if main_window and hasattr(main_window, 'raw_points') and main_window.raw_points:
            scaled = main_window.scale_points(main_window.raw_points)
            self.set_points(scaled)
        super().resizeEvent(event)

class MainWindow(QMainWindow):
    """
    主窗口，包含按钮、图形区和主要功能
    """
    def __init__(self):
        super().__init__()
        self.setWindowTitle("铁匠Excel坐标转DXF工具")
        #这里可以修改窗口大小
        self.setGeometry(100, 100, 400, 400)

        # 设置窗口和任务栏图标（兼容打包和源码运行）
        #如果要修改图标，把图标文件放在同级目录下，命名为"doro.ico"，当然你也可以改代码
        if getattr(sys, 'frozen', False):
            icon_path = os.path.join(sys._MEIPASS, "doro.ico")
        else:
            icon_path = "doro.ico"
        self.setWindowIcon(QIcon(icon_path))

        # 主界面布局
        main_widget = QWidget()
        main_layout = QVBoxLayout(main_widget)

        # 选择Excel文件按钮
        self.button = QPushButton("选择Excel文件")
        self.button.clicked.connect(self.open_file)

        # 导出DXF按钮
        self.export_button = QPushButton("导出为DXF")
        self.export_button.clicked.connect(self.export_dxf)

        # 图形预览区
        self.graph_widget = GraphWidget(self)

        # 添加控件到布局
        main_layout.addWidget(self.button)
        main_layout.addWidget(self.export_button)
        main_layout.addWidget(self.graph_widget)
        self.setCentralWidget(main_widget)

        # 保存原始坐标点
        self.raw_points = []

    def open_file(self):
        """
        打开文件对话框，读取Excel坐标并显示预览
        """
        file_path, _ = QFileDialog.getOpenFileName(self, "选择Excel文件", "", "Excel Files (*.xlsx *.xls)")
        if file_path:
            points = self.read_points_from_excel(file_path)
            self.raw_points = points  # 保存原始坐标
            scaled_points = self.scale_points(points)
            self.graph_widget.set_points(scaled_points)

    def read_points_from_excel(self, file_path):
        """
        读取Excel文件中的坐标点，假设A列为X，B列为Y，跳过表头
        """
        wb = openpyxl.load_workbook(file_path)
        ws = wb.active
        points = []
        for row in ws.iter_rows(min_row=2, values_only=True):
            x, y = row[0], row[1]
            if x is not None and y is not None:
                points.append((float(x), float(y)))
        return points

    def scale_points(self, points):
        """
        将原始坐标缩放到窗口大小，保持比例、居中，并翻转Y轴（适应屏幕坐标系）
        """
        if not points:
            return []
        xs, ys = zip(*points)
        min_x, max_x = min(xs), max(xs)
        min_y, max_y = min(ys), max(ys)
        w, h = self.graph_widget.width(), self.graph_widget.height()
        margin = 40  # 边距
        dx = max_x - min_x if max_x != min_x else 1
        dy = max_y - min_y if max_y != min_y else 1
        scale = min((w - margin * 2) / dx, (h - margin * 2) / dy)
        offset_x = (w - scale * dx) / 2
        offset_y = (h - scale * dy) / 2
        # y轴翻转
        scaled = [(
            int((x - min_x) * scale + offset_x),
            int(h - ((y - min_y) * scale + offset_y))
        ) for x, y in points]
        return scaled

    def export_dxf(self):
        """
        导出DXF文件，包含法兰盘和坐标连线
        """
        import ezdxf
        from PyQt5.QtWidgets import QMessageBox

        file_path, _ = QFileDialog.getSaveFileName(self, "保存DXF文件", "", "DXF Files (*.dxf)")
        if not file_path:
            return

        doc = ezdxf.new(dxfversion="R2010")
        msp = doc.modelspace()

        # 统一颜色
        color = 7  # 黑色/白色

        # 1. 画法兰盘（中心大圆和四个小孔）
        msp.add_circle((0, 0), 6, dxfattribs={"color": color})  # 半径6，直径12
        for dx, dy in [(10, 0), (-10, 0), (0, 10), (0, -10)]:
            msp.add_circle((dx, dy), 1.5, dxfattribs={"color": color})  # 半径1.5，直径3

        # 2. 画坐标连线
        if hasattr(self, "raw_points") and self.raw_points and len(self.raw_points) > 1:
            for i in range(len(self.raw_points) - 1):
                x1, y1 = self.raw_points[i]
                x2, y2 = self.raw_points[i + 1]
                msp.add_line((x1, y1), (x2, y2), dxfattribs={"color": color})

        try:
            doc.saveas(file_path)
            QMessageBox.information(self, "导出成功", f"DXF文件已保存到：\n{file_path}")
        except Exception as e:
            QMessageBox.warning(self, "导出失败", str(e))

if __name__ == "__main__":
    # 程序入口
    app = QApplication(sys.argv)
    window = MainWindow()
    window.show()
    sys.exit(app.exec_())
