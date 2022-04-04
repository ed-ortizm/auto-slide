from configparser import ConfigParser, ExtendedInterpolation
import glob
import time

import numpy as np

from pptx import Presentation
from pptx.enum.shapes import MSO_SHAPE
from pptx.util import Inches
from pptx.enum.dml import MSO_THEME_COLOR
from pptx.dml.color import RGBColor
###############################################################################
start_time = time.time()
###############################################################################
parser = ConfigParser(interpolation=ExtendedInterpolation())
config_file_name = "generate_slides.ini"
parser.read(f"{config_file_name}")

###############################################################################
presentation_name = parser.get("file", "presentation_name")
presentation = Presentation()

TITLE_AND_CONTENT = 1
BLANK = 6
###############################################################################
print(f"Set title", end="\n")

layout = presentation.slide_layouts[BLANK]
slide = presentation.slides.add_slide(layout)
###############################################################################
shapes = slide.shapes

left = Inches(4.0)
top = Inches(4.0)
width = Inches(3.0)
height = Inches(3.0)

shape = shapes.add_shape(
    MSO_SHAPE.ROUNDED_RECTANGLE,
    left,
    top,
    width,
    height
)
fill = shape.fill
fill.solid()
fill.fore_color.rgb = RGBColor(255, 0, 0)

###############################################################################
line = shape.line
line.color.rgb = RGBColor(255, 0, 0)
line.color.brightness = 0.5  # 50% lighter
# line.width = Pt(2.5)
# for number in range(10):
#
#     layout = presentation.slide_layouts[number]
#     slide = presentation.slides.add_slide(layout)
###############################################################################
save_to = parser.get("directory", "save_to")
presentation.save(f"{save_to}/{presentation_name}.pptx")

###############################################################################
finish_time = time.time()
print(f"Run time: {finish_time - start_time: .2f}", end="\n")
