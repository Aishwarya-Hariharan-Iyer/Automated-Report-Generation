import openpyxl
from pptx import Presentation
from openpyxl.chart import BarChart3D,AreaChart, Reference
import pandas as pd
import matplotlib.pyplot as plt
from pptx.util import Inches
import matplotlib.ticker as mtick
from PIL import Image

def getTitle(dataframe):
    # obtain the visualization title from excel
    thisTitle = dataframe.iloc[1]['Visualisation Title']
    return thisTitle

def getSlideNo(dataframe):
    # obtain slide number from excel to set as slide name
    slideNo = str(dataframe.iloc[1]['Slide number'])
    return slideNo

def getMonthlyData(dataframe, y1, y2):
    #get the data-containing columns from dataframe
    df2 = dataframe.iloc[:, y1: y2]
    return df2

def add_value_labels(ax, typ, title="", spacing=5):
    space = spacing
    va = 'bottom'

    if typ == 'bar':
        for i in ax.patches:
            y_value = i.get_height()
            x_value = i.get_x() + i.get_width() / 2

            if title=="PMI Trend":
                label = "{:.1f}".format(y_value)
                ax.annotate(label,(x_value, y_value), xytext=(0, space),
                            textcoords="offset points", ha='center', va=va, fontsize = 5)
            else:
                label = "{:.0f}".format(y_value)
                ax.annotate(label, (x_value, y_value), xytext=(0, space),
                            textcoords="offset points", ha='center', va=va, fontsize=5)
    if typ == 'line':
        for line in ax.lines:
            for x_value, y_value in zip(line.get_xdata(), line.get_ydata()):
                label = "{:.1f}".format(y_value)
                ax.annotate(label,(x_value, y_value), xytext=(0, space),
                    textcoords="offset points", ha='center', va=va, fontsize=5)

def isfloat(num):
    try:
        float(num)
        return True
    except ValueError:
        return False

def px_to_inches(path):
    #function to center given image
    im = Image.open(path)
    width = im.width / im.info['dpi'][0]
    height = im.height / im.info['dpi'][1]
    return (width, height)



def findDimensions(slide_size, imgD):
    # standardize picture dimensions
    left = Inches(slide_size[0] - imgD[0]) / 2
    top = Inches(slide_size[1] - imgD[1]) / 2
    return (left, top)

def fixData(df):
    #extract the column headers and restrict dataset to numerical values while plotting
    df.columns = df.iloc[0]
    df = df[1:]
    return df

def makeBlankSlide(prs):
    #generate empty slides
    slide_layout = prs.slide_layouts[6]
    slide = prs.slides.add_slide(slide_layout)