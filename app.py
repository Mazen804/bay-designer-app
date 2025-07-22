import streamlit as st
import matplotlib.pyplot as plt
import matplotlib.patches as patches
import io
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import MSO_ANCHOR, MSO_AUTO_SIZE, PP_ALIGN
from pptx.enum.shapes import MSO_SHAPE

# --- Page Configuration ---
st.set_page_config(layout="wide", page_title="Storage Bay Designer")

# --- Main Title ---
st.title("Storage Bay Designer")
st.markdown("Use the sidebar to manage and configure your bay groups. The design will update live in the main area.")

# --- Helper Functions ---

def draw_dimension_line(ax, x1, y1, x2, y2, text, is_vertical=False, offset=10, color='black', fontsize=9):
    """Draws a dimension line with arrows and text on the matplotlib axis."""
    ax.plot([x1, x2], [y1, y2], color=color, lw=1)
    if is_vertical:
        ax.plot(x1, y1, marker='v', color=color, markersize=5)
        ax.plot(x2, y2, marker='^', color=color, markersize=5)
        ax.text(x1 + offset, (y1 + y2) / 2, text, va='center', ha='left', fontsize=fontsize, rotation=90, color=color)
    else:
        ax.plot(x1, y1, marker='<', color=color, markersize=5)
        ax.plot(x2, y2, marker='>', color=color, markersize=5)
        ax.text((x1 + x2) / 2, y1 + offset, text, va='bottom', ha='center', fontsize=fontsize, color=color)

def draw_bay_group(params):
    """Main function to draw a group of bays using Matplotlib for the LIVE PREVIEW."""
    # Unpack parameters
    num_bays = params['num_bays']
    bay_width = par