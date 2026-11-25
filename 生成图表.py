#!/usr/bin/env python3
"""从统计结果.json生成报告图。依赖 reportlab；PNG 转换使用 bundled Node/sharp。

运行：python 生成图表.py
默认读本目录“数据结果/统计结果.json”，输出“图表”中的 SVG 与 2 倍 PNG。
可通过 CODEX_CHART_NODE / CODEX_CHART_SHARP 指定 Node 可执行文件和 sharp 模块目录。
所有统计值均读取分析结果，不在绘图时重新估计或人为调整。
"""
from pathlib import Path
import json
import os
import subprocess
from reportlab.graphics.shapes import Drawing, Rect, Line, String, Circle
from reportlab.graphics import renderSVG
from reportlab.lib.colors import HexColor
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont

BASE=Path(__file__).resolve().parent
DATA=json.loads((BASE/'数据结果'/'统计结果.json').read_text(encoding='utf-8'))
OUT=BASE/'图表'
OUT.mkdir(parents=True,exist_ok=True)
RUNTIME=Path.home()/'.cache/codex-runtimes/codex-primary-runtime/dependencies/node'
NODE=Path(os.environ.get('CODEX_CHART_NODE',str(RUNTIME/'bin/node')))
SHARP=Path(os.environ.get('CODEX_CHART_SHARP',str(RUNTIME/'node_modules/sharp')))
FONT_FILE=Path(os.environ.get('CODEX_CHART_FONT','/System/Library/Fonts/Supplemental/Arial Unicode.ttf'))
# 为中文字符串使用正确的全角字宽，避免右对齐标签侵入绘图区。
pdfmetrics.registerFont(TTFont('ChartRegular',str(FONT_FILE)))
pdfmetrics.registerFont(TTFont('ChartBold',str(FONT_FILE)))

W,H=1000,500
BLACK='#17212C'
GRAY='#526171'
GRID='#E3E8ED'
BLUE='#2463A8'
ORANGE='#C96822'
WHITE='#FFFFFF'

def text(d,x,y,s,size=20,color=BLACK,anchor='start',bold=False):
    d.add(String(x,y,str(s),fontName='ChartBold' if bold else 'ChartRegular',
                 fontSize=size,fillColor=HexColor(color),textAnchor=anchor))
def line(d,x1,y1,x2,y2,color=BLACK,width=1):
    d.add(Line(x1,y1,x2,y2,strokeColor=HexColor(color),strokeWidth=width))
def rect(d,x,y,w,h,color):
    d.add(Rect(x,y,w,h,strokeColor=None,fillColor=HexColor(color)))
def circle(d,x,y,r,color):
    d.add(Circle(x,y,r,strokeColor=HexColor(WHITE),strokeWidth=1.5,fillColor=HexColor(color)))
def canvas(title,subtitle):
    d=Drawing(W,H)
    rect(d,0,0,W,H,WHITE)
    text(d,38,460,title,28,bold=True)
    text(d,38,428,subtitle,18,GRAY)
    return d
def footer(d,first,second=None):
    line(d,38,53,962,53,GRID,1)
    text(d,38,30,first,15,GRAY)
    if second:text(d,38,10,second,15,GRAY)
def vert_ci(d,x,lo,hi,color,cap=9):
    line(d,x,lo,x,hi,color,2.8)
    line(d,x-cap,lo,x+cap,lo,color,2.8)
    line(d,x-cap,hi,x+cap,hi,color,2.8)
def hor_ci(d,y,lo,hi,color,cap=9):
    line(d,lo,y,hi,y,color,3.4)
    line(d,lo,y-cap,lo,y+cap,color,3.4)
    line(d,hi,y-cap,hi,y+cap,color,3.4)
def save(d,name):
    svg=renderSVG.drawToString(d)
    # SVG 原生文本保留可编辑性；librsvg 使用系统中文字体绘制 PNG。
    svg=svg.replace('font-family: ChartBold;',"font-family: 'PingFang SC', 'Heiti SC', 'Hiragino Sans GB', sans-serif; font-weight: 700;")
    svg=svg.replace('font-family: ChartRegular;',"font-family: 'PingFang SC', 'Heiti SC', 'Hiragino Sans GB', sans-serif;")
    svg=svg.replace('<title>...</title>',f'<title>{name}</title>')
    svgpath=OUT/f'{name}.svg'; pngpath=OUT/f'{name}.png'
    svgpath.write_text(svg,encoding='utf-8')
    js="const sharp=require(process.argv[1]);sharp(process.argv[2],{density:144}).png().toFile(process.argv[3]).catch(e=>{console.error(e);process.exit(1)});"
    subprocess.run([str(NODE),'-e',js,str(SHARP),str(svgpath),str(pngpath)],check=True)
    print(f'{name}: {svgpath.name}, {pngpath.name}')

def stars():
    n=DATA['overall']['n']; low=DATA['overall']['rate']
    d=canvas('评论星级分布',f'共 {n:,} 条评论；低评分（1—2 星）占 {low:.1%}')
    left,right,bottom,top=105,938,95,374
    for v in range(0,501,100):
        y=bottom+v/500*(top-bottom)
        line(d,left,y,right,y,GRID)
        text(d,left-15,y-6,str(v),17,GRAY,'end')
    text(d,left-12,398,'评论数（条）',17,GRAY)
    line(d,left,bottom,right,bottom,GRAY)
    for i,item in enumerate(DATA['stars']):
        x=190+i*166; h=item['count']/500*(top-bottom)
        rect(d,x-45,bottom,90,h,ORANGE if item['star']<=2 else BLUE)
        text(d,x,bottom+h+12,f"{item['count']} 条 · {item['share']:.1%}",19,BLACK,'middle')
        text(d,x,bottom-30,f"{item['star']} 星",20,BLACK,'middle')
    footer(d,'来源：卖家评论.xlsx；统计单位为评论记录。橙色表示本研究定义的低评分。')
    save(d,'01_星级分布')

def image_compare():
    s=DATA['main']; diff=s['difference']*100; ci=[x*100 for x in s['difference_newcombe95']]
    d=canvas('图片标记与低评分比例',f'两组相差 {diff:.2f} 个百分点（差值 95% 区间：{ci[0]:.2f}—{ci[1]:.2f}）')
    left,right,bottom,top=265,930,110,360
    for v in range(0,101,20):
        x=left+v/100*(right-left)
        line(d,x,bottom,x,top,GRID)
        text(d,x,bottom-28,f'{v}%',18,GRAY,'middle')
    line(d,left,bottom,right,bottom,GRAY)
    for label,key,row,y,color in [('有图片标记','image',0,305,ORANGE),('未标记图片','unmarked',1,183,BLUE)]:
        value=s[f'{key}_rate']; low,high=s[f'{key}_wilson95']
        xx=lambda a:left+a*(right-left)
        k=sum(s['table'][row]); count=s['table'][row][0]
        text(d,42,y+4,label,24,bold=True)
        text(d,42,y-25,f'{count} / {k:,} 条低评分',18,GRAY)
        hor_ci(d,y,xx(low),xx(high),color)
        circle(d,xx(value),y,7,color)
        text(d,xx(value),y+27,f'{value:.2%}',24,BLACK,'middle',True)
        text(d,xx(value),y-36,f'95% 区间 {low:.2%}—{high:.2%}',17,GRAY,'middle')
    footer(d,'点为低评分比例，横线为 Wilson 95% 区间；低评分为 1—2 星。',
           '空白仅记为“未标记图片”；区间基于独立观测假设，关联不代表因果。')
    save(d,'02_图片标记与低评分')

def monthly():
    d=canvas('分月比较图片标记与低评分比例','各月方向一致；首尾月时间覆盖不完整，9 月图片标记组仅 6 条')
    left,right,bottom,top=95,935,126,376
    for v in range(0,101,20):
        y=bottom+v/100*(top-bottom)
        line(d,left,y,right,y,GRID)
        text(d,left-14,y-6,f'{v}%',17,GRAY,'end')
    line(d,left,bottom,right,bottom,GRAY)
    circle(d,640,400,5,ORANGE);text(d,654,394,'有图片标记',17)
    circle(d,812,400,5,BLUE);text(d,826,394,'未标记图片',17)
    for i,(month,s) in enumerate(DATA['monthly'].items()):
        center=194+i*219
        for key,delta,color in [('image',-27,ORANGE),('unmarked',27,BLUE)]:
            x=center+delta;val=s[f'{key}_rate']; lo,hi=s[f'{key}_wilson95']
            yy=lambda p:bottom+p*(top-bottom)
            vert_ci(d,x,yy(lo),yy(hi),color,8)
            circle(d,x,yy(val),6,color)
            # 标签朝两点外侧排布，避免遮住同月另一组区间。
            text(d,x+(-12 if delta<0 else 12),yy(val)-5,f'{val:.1%}',17,BLACK,'end' if delta<0 else 'start')
        text(d,center,95,f'{int(month[-2:])} 月',21,BLACK,'middle')
        ns=[sum(row) for row in s['table']]
        text(d,center,70,f'n = {ns[0]} / {ns[1]}',17,GRAY,'middle')
    footer(d,'点与竖线：低评分比例及 Wilson 95% 区间；n 顺序为有图片标记 / 未标记图片。',
           '数据覆盖 2026-06-02—2026-09-11；分月结果用于探索性稳健性检查。')
    save(d,'03_分月比较')

def themes():
    d=canvas('低评分评论中的预设主题命中率','分母为 407 条低评分评论；结果来自五条预设的负向关键词规则')
    left,right,bottom,top=242,840,95,384
    for v in range(0,31,5):
        x=left+v/30*(right-left)
        line(d,x,bottom,x,top,GRID)
        text(d,x,bottom-25,f'{v}%',17,GRAY,'middle')
    line(d,left,bottom,right,bottom,GRAY)
    for i,item in enumerate(DATA['negative_themes']):
        y=355-i*53;w=item['share']/.30*(right-left)
        text(d,left-19,y-6,item['topic'],21,BLACK,'end')
        rect(d,left,y-15,w,30,BLUE)
        text(d,left+w+13,y-6,f"{item['share']:.2%}（{item['count']} 条）",19)
    footer(d,'一条评论可命中多个主题；五类合计不代表互斥构成，未命中不代表没有问题。',
           '规则命中尚未经过独立人工验证；命中标签不能直接视为已确认的投诉原因。')
    save(d,'04_低星主题')

if __name__=='__main__':
    stars(); image_compare(); monthly(); themes()
