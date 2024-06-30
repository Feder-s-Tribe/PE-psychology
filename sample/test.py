import plotly.graph_objects as go
import os

# 创建一个简单的柱状图
fig = go.Figure(data=[go.Bar(x=['A', 'B', 'C'], y=[1, 3, 2])])

# 尝试将图像保存为JPEG格式
os.environ["KaleidoScope"] = "graph_objects"
fig.write_image("test_image.jpeg", engine="kaleido")