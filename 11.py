from paddleocr import PaddleOCR, draw_ocr
from PIL import Image
import matplotlib.pyplot as plt
 
# 1. 初始化 PaddleOCR 引擎
# use_angle_cls=True 表示加载方向分类器，自动纠正图片方向
# lang="ch" 表示识别中文 (支持 ch, en, fr, german, korean, japan)
ocr = PaddleOCR(use_angle_cls=True, lang="ch") 
 
# 2. 定义图片路径 (可以是本地路径，也可以是网络图片的 Numpy 数组)
img_path = r'C:\Users\Admin\Desktop\Snipaste_2026-02-06_10-50-16.png'
 
# 为了演示，如果没有 test.jpg，我们提示用户
import os
if not os.path.exists(img_path):
    print(f"错误：请在当前目录下放一张名为 {img_path} 的带有文字的图片进行测试！")
    exit()
 
# 3. 开始推理
# result 是一个列表，列表中的每一项对应一行文字
# 结构：[ [ [[x1,y1],[x2,y2],[x3,y3],[x4,y4]], (text, score) ], ... ]
result = ocr.ocr(img_path, cls=True)
 
# 4. 打印结果
print("------------------ 识别结果 ------------------")
# ocr.ocr 返回的结果是列表的列表 (PaddleOCR v2.6+ 结构略有变化，通常 result[0] 才是内容)
for idx in range(len(result)):
    res = result[idx]
    if res is None: # 未识别到内容
        print(f"未在 {img_path} 中识别到文字")
        continue
        
    for line in res:
        print(line)
 
# 5. (可选) 结果可视化
# 提取坐标和文字
boxes = [line[0] for line in result[0]]
txts = [line[1][0] for line in result[0]]
scores = [line[1][1] for line in result[0]]
 
# 使用 Pillow 加载图片
image = Image.open(img_path).convert('RGB')
# 使用 PaddleOCR 自带的工具画框
im_show = draw_ocr(image, boxes, txts, scores, font_path='./fonts/simfang.ttf') # 注意：需要提供一个支持中文的字体文件路径
im_show = Image.fromarray(im_show)
im_show.save('result.jpg')
print("---------------------------------------------")
print("可视化结果已保存为 result.jpg")
