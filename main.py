
import pandas as pd
import  math
import docx
from docx.shared import Pt
from docx.enum.table import WD_ROW_HEIGHT, WD_ALIGN_VERTICAL
from docx.enum.text import WD_PARAGRAPH_ALIGNMENT
from utils import set_cell_margins, set_cell_border
from bakoverall import get_overall_table

excel_path = 'CCWFO-10-23_21_13.xlsx'
doc = docx.Document()
table2 = get_overall_table(excel_path,doc)

doc.save("table2.docx")