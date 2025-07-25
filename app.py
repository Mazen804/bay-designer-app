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

def hex_to_rgb(hex_color):
    """Converts a hex color string to an RGB tuple."""
    hex_color = hex_color.lstrip('#')
    return tuple(int(hex_color[i:i+2], 16) for i in (0, 2, 4))

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
    bay_width = params['bay_width']
    total_height = params['total_height']
    ground_clearance = params['ground_clearance']
    shelf_thickness = params['shelf_thickness']
    side_panel_thickness = params['side_panel_thickness']
    num_cols = params['num_cols']
    num_rows = params['num_rows']
    has_top_cap = params['has_top_cap']
    color = params['color']
    bin_heights = params['bin_heights']
    zoom_factor = params.get('zoom', 1.0)
    bin_split_thickness = shelf_thickness

    # --- Calculations ---
    core_width = num_bays * bay_width
    total_group_width = core_width + (2 * side_panel_thickness)
    
    fig, ax = plt.subplots(figsize=(12, 12))

    # --- Draw Structure using a stable coordinate system ---
    # The core bays are drawn from x=0 to x=core_width. Side panels are drawn in the negative and positive space around this.
    structure_height = total_height - ground_clearance
    
    # Left side panel
    ax.add_patch(patches.Rectangle((-side_panel_thickness, 0), side_panel_thickness, total_height, facecolor=color))
    # Right side panel
    ax.add_patch(patches.Rectangle((core_width, 0), side_panel_thickness, total_height, facecolor=color))

    current_x = 0
    for bay_idx in range(num_bays):
        net_width_per_bay = bay_width
        total_internal_dividers = (num_cols - 1) * bin_split_thickness
        bin_width = (net_width_per_bay - total_internal_dividers) / num_cols if num_cols > 0 else 0

        bin_start_x = current_x
        if num_cols > 1:
            for i in range(1, num_cols):
                split_x = bin_start_x + (i * bin_width) + ((i-1) * bin_split_thickness)
                ax.add_patch(patches.Rectangle((split_x, ground_clearance), bin_split_thickness, structure_height, facecolor=color))
        
        if bay_idx < num_bays - 1:
             divider_x = current_x + bay_width
             ax.plot([divider_x, divider_x], [ground_clearance, structure_height], color='#aaaaaa', lw=1, linestyle='--')

        current_x += bay_width

    # --- Draw Horizontal Shelves & Bin Height Dimensions ---
    current_y = ground_clearance
    dim_offset_x = 0.05 * core_width
    pitch_offset_x = dim_offset_x * 2.5

    for i in range(num_rows):
        shelf_bottom_y = current_y
        ax.add_patch(patches.Rectangle((-side_panel_thickness, shelf_bottom_y), total_group_width, shelf_thickness, facecolor=color))
        shelf_top_y = shelf_bottom_y + shelf_thickness
        
        if i < len(bin_heights):
            net_bin_h = bin_heights[i]
            pitch_h = net_bin_h + shelf_thickness
            level_name = chr(65 + i)
            
            bin_bottom_y = shelf_top_y
            bin_top_y = bin_bottom_y + net_bin_h
            draw_dimension_line(ax, core_width + side_panel_thickness + dim_offset_x, bin_bottom_y, core_width + side_panel_thickness + dim_offset_x, bin_top_y, f"{net_bin_h:.1f}", is_vertical=True, offset=5, color='#3b82f6')
            
            pitch_top_y = shelf_bottom_y + pitch_h
            draw_dimension_line(ax, core_width + side_panel_thickness + pitch_offset_x, shelf_bottom_y, core_width + side_panel_thickness + pitch_offset_x, pitch_top_y, f"{pitch_h:.1f}", is_vertical=True, offset=5, color='black')

            ax.text(-side_panel_thickness - dim_offset_x, (bin_bottom_y + bin_top_y) / 2, level_name, va='center', ha='center', fontsize=12, fontweight='bold')
            
            current_y = bin_top_y

    if has_top_cap:
        ax.add_patch(patches.Rectangle((-side_panel_thickness, total_height - shelf_thickness), total_group_width, shelf_thickness, facecolor=color))

    # --- Draw Main Dimension Lines ---
    dim_offset_y = 0.05 * total_height
    draw_dimension_line(ax, -side_panel_thickness, -dim_offset_y * 2, core_width + side_panel_thickness, -dim_offset_y * 2, f"Total Group Width: {total_group_width:.0f} mm", offset=10)
    draw_dimension_line(ax, -side_panel_thickness - (dim_offset_x * 4), 0, -side_panel_thickness - (dim_offset_x * 4), total_height, f"Total Height: {total_height:.0f} mm", is_vertical=True, offset=10)

    # --- Draw Bin Width Dimensions above the bay ---
    if num_cols > 0:
        dim_y_pos = total_height + dim_offset_y
        loop_current_x = 0
        for bay_idx in range(num_bays):
            net_width_per_bay = bay_width
            total_internal_dividers = (num_cols - 1) * bin_split_thickness
            bin_width = (net_width_per_bay - total_internal_dividers) / num_cols if num_cols > 0 else 0
            
            bin_start_x = loop_current_x
            for i in range(num_cols):
                dim_start_x = bin_start_x + (i * (bin_width + bin_split_thickness))
                dim_end_x = dim_start_x + bin_width
                draw_dimension_line(ax, dim_start_x, dim_y_pos, dim_end_x, dim_y_pos, f"{bin_width:.1f}", offset=10, color='#3b82f6')
            
            loop_current_x += bay_width

    # --- Final Touches ---
    ax.set_aspect('equal', adjustable='box')
    padding_x = core_width * 0.4 + side_panel_thickness
    ax.set_xlim((-padding_x) * zoom_factor, (core_width + padding_x) * zoom_factor)
    ax.set_ylim(-dim_offset_y * 3 * zoom_factor, total_height + dim_offset_y * 2 * zoom_factor)
    ax.axis('off')
    
    return fig

def create_editable_powerpoint(bay_groups):
    """Creates a PowerPoint presentation from bay group data using native shapes."""
    prs = Presentation()
    
    for group_data in bay_groups:
        slide = prs.slides.add_slide(prs.slide_layouts[6]) # Blank layout
        
        title_shape = slide.shapes.add_textbox(Inches(0.5), Inches(0.2), Inches(9), Inches(0.5))
        title_shape.text = f"Design for: {group_data['name']}"

        # --- Unpack Parameters ---
        num_bays, bay_width, total_height, ground_clearance, shelf_thickness, side_panel_thickness, num_cols, num_rows, has_top_cap, color_hex, bin_heights = (
            group_data['num_bays'], group_data['bay_width'], group_data['total_height'],
            group_data['ground_clearance'], group_data['shelf_thickness'], group_data['side_panel_thickness'],
            group_data['num_cols'], group_data['num_rows'], group_data['has_top_cap'],
            group_data['color'], group_data['bin_heights']
        )
        bin_split_thickness = shelf_thickness

        # --- Define Drawing Area and Scale on Slide ---
        canvas_left, canvas_top, canvas_width, canvas_height = Inches(1.5), Inches(1), Inches(7), Inches(5.5)
        total_group_width = (num_bays * bay_width) + (2 * side_panel_thickness)
        scale = min(canvas_width / (total_group_width + 400), canvas_height / (total_height + 200))

        def pt_to_emu(points):
            return int(points * 12700)

        def add_shape(left_mm, top_mm, width_mm, height_mm, color_hex):
            left = canvas_left + left_mm * scale
            top = canvas_top + (total_height - top_mm - height_mm) * scale
            width = width_mm * scale
            height = height_mm * scale
            shape = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, left, top, width, height)
            shape.fill.solid()
            shape.fill.fore_color.rgb = RGBColor(*hex_to_rgb(color_hex))
            shape.line.fill.background()
            return shape
        
        def add_dimension(start_x, start_y, end_x, end_y, text, is_vertical=False):
            line = slide.shapes.add_connector(1, start_x, start_y, end_x, end_y)
            line.line.fill.solid()
            line.line.fill.fore_color.rgb = RGBColor(0,0,0)
            line.line.begin_arrow_type = 2
            line.line.end_arrow_type = 2

            if is_vertical:
                text_left = start_x + pt_to_emu(5)
                text_top = start_y + (end_y - start_y) / 2 - pt_to_emu(20)
                textbox = slide.shapes.add_textbox(text_left, text_top, Inches(0.5), Inches(0.5))
                textbox.rotation = 270.0
            else:
                text_left = start_x + (end_x - start_x) / 2 - pt_to_emu(20)
                text_top = start_y - pt_to_emu(12)
                textbox = slide.shapes.add_textbox(text_left, text_top, Inches(0.5), Inches(0.5))
            
            p = textbox.text_frame.paragraphs[0]
            p.text = text
            p.font.size = Pt(8)
            p.alignment = PP_ALIGN.CENTER
            textbox.text_frame.auto_size = MSO_AUTO_SIZE.SHAPE_TO_FIT_TEXT

        # --- Draw Structure using PPTX Shapes ---
        structure_height = total_height - ground_clearance
        add_shape(0, 0, side_panel_thickness, total_height, color_hex)
        current_x_mm = side_panel_thickness

        for bay_idx in range(num_bays):
            net_width_per_bay = bay_width
            total_internal_dividers = (num_cols - 1) * bin_split_thickness
            bin_width = (net_width_per_bay - total_internal_dividers) / num_cols if num_cols > 0 else 0

            bin_start_x_mm = current_x_mm
            if num_cols > 1:
                for i in range(1, num_cols):
                    split_x_mm = bin_start_x_mm + (i * bin_width) + ((i-1) * bin_split_thickness)
                    add_shape(split_x_mm, ground_clearance, bin_split_thickness, structure_height, color_hex)
            
            for i in range(num_cols):
                dim_start_x = canvas_left + (bin_start_x_mm + i * (bin_width + bin_split_thickness)) * scale
                dim_end_x = dim_start_x + (bin_width * scale)
                dim_y = canvas_top - pt_to_emu(20)
                add_dimension(dim_start_x, dim_y, dim_end_x, dim_y, f"{bin_width:.1f}")

            current_x_mm += bay_width

        add_shape(current_x_mm, 0, side_panel_thickness, total_height, color_hex)

        current_y_mm = ground_clearance
        for i in range(num_rows):
            shelf_bottom_y = current_y_mm
            add_shape(0, shelf_bottom_y, total_group_width, shelf_thickness, color_hex)
            shelf_top_y = shelf_bottom_y + shelf_thickness
            
            if i < len(bin_heights):
                net_bin_h = bin_heights[i]
                pitch_h = net_bin_h + shelf_thickness

                dim_start_y = canvas_top + (total_height - (shelf_top_y + net_bin_h)) * scale
                dim_end_y = canvas_top + (total_height - shelf_top_y) * scale
                dim_x = canvas_left + (total_group_width + 50) * scale
                add_dimension(dim_x, dim_start_y, dim_x, dim_end_y, f"{net_bin_h:.1f}", is_vertical=True)
                
                pitch_dim_start_y = canvas_top + (total_height - (shelf_bottom_y + pitch_h)) * scale
                pitch_dim_end_y = canvas_top + (total_height - shelf_bottom_y) * scale
                pitch_dim_x = canvas_left + (total_group_width + 150) * scale
                add_dimension(pitch_dim_x, pitch_dim_start_y, pitch_dim_x, pitch_dim_end_y, f"{pitch_h:.1f}", is_vertical=True)

                current_y_mm += shelf_thickness + net_bin_h

        if has_top_cap:
            add_shape(0, total_height - shelf_thickness, total_group_width, shelf_thickness, color_hex)

        total_w_y = canvas_top + canvas_height + pt_to_emu(20)
        add_dimension(canvas_left, total_w_y, canvas_left + total_group_width * scale, total_w_y, f"Total Width: {total_group_width:.0f} mm")
        
        total_h_x = canvas_left - pt_to_emu(40)
        add_dimension(total_h_x, canvas_top, total_h_x, canvas_top + total_height * scale, f"Total Height: {total_height:.0f} mm", is_vertical=True)

    ppt_buf = io.BytesIO()
    prs.save(ppt_buf)
    ppt_buf.seek(0)
    return ppt_buf

# --- Initialize Session State ---
if 'bay_groups' not in st.session_state:
    st.session_state.bay_groups = [{
        "name": "Group A", "num_bays": 2, "bay_width": 1050.0, "total_height": 2000.0,
        "ground_clearance": 50.0, "shelf_thickness": 18.0, "side_panel_thickness": 18.0,
        "num_cols": 4, "num_rows": 5, "has_top_cap": True, "color": "#4A90E2",
        "bin_heights": [350.0] * 5,
        "zoom": 1.0
    }]

# --- Sidebar Controls ---
st.sidebar.header("Manage Bay Groups")

with st.sidebar.form("new_group_form"):
    new_group_name = st.text_input("New Group Name", "New Group")
    add_group_submitted = st.form_submit_button("Add Group")
    if add_group_submitted:
        new_group = st.session_state.bay_groups[0].copy()
        new_group['name'] = new_group_name
        st.session_state.bay_groups.append(new_group)
        st.rerun()

if len(st.session_state.bay_groups) > 1:
    if st.sidebar.button("Remove Last Group"):
        st.session_state.bay_groups.pop()
        st.rerun()

st.sidebar.markdown("---")

group_names = [g['name'] for g in st.session_state.bay_groups]
selected_group_name = st.sidebar.selectbox("Select Group to Edit", group_names)
active_group_idx = group_names.index(selected_group_name)
group_data = st.session_state.bay_groups[active_group_idx]

st.sidebar.header(f"Configuration for: {group_data['name']}")

# --- Dynamic Height Calculation Callbacks ---
def distribute_total_height():
    active_group = st.session_state.bay_groups[active_group_idx]
    num_shelves_for_calc = active_group['num_rows'] + (1 if active_group['has_top_cap'] else 0)
    total_shelf_thickness = num_shelves_for_calc * active_group['shelf_thickness']
    available_space = active_group['total_height'] - active_group['ground_clearance'] - total_shelf_thickness
    
    if available_space > 0 and active_group['num_rows'] > 0:
        uniform_net_h = available_space / active_group['num_rows']
        active_group['bin_heights'] = [uniform_net_h] * active_group['num_rows']
        for j in range(active_group['num_rows']):
            st.session_state[f"level_{active_group_idx}_{j}"] = uniform_net_h

# --- Configuration Inputs ---
st.sidebar.subheader("Structure")
group_data['num_bays'] = st.sidebar.number_input("Number of Bays in Group", min_value=1, value=int(group_data['num_bays']), key=f"num_bays_{active_group_idx}")
group_data['bay_width'] = st.sidebar.number_input("Width per Bay (mm)", min_value=1.0, value=float(group_data['bay_width']), key=f"bay_width_{active_group_idx}")
group_data['total_height'] = st.sidebar.number_input("Target Total Height (mm)", min_value=1.0, value=float(group_data['total_height']), key=f"total_height_{active_group_idx}", on_change=distribute_total_height, help="Set this to automatically distribute height among bins.")
group_data['ground_clearance'] = st.sidebar.number_input("Ground Clearance (mm)", min_value=0.0, value=float(group_data['ground_clearance']), key=f"ground_clearance_{active_group_idx}")
group_data['has_top_cap'] = st.sidebar.checkbox("Add Top Cap", value=group_data['has_top_cap'], key=f"has_top_cap_{active_group_idx}")

st.sidebar.subheader("Layout")
group_data['num_rows'] = st.sidebar.number_input("Shelves (Rows)", min_value=1, value=int(group_data['num_rows']), key=f"num_rows_{active_group_idx}")
group_data['num_cols'] = st.sidebar.number_input("Bin-Split (Columns)", min_value=1, value=int(group_data['num_cols']), key=f"num_cols_{active_group_idx}")

st.sidebar.markdown("**Individual Net Bin Heights**")
if len(group_data['bin_heights']) != group_data['num_rows']:
    distribute_total_height()

current_bin_heights = []
for j in range(group_data['num_rows']):
    level_name = chr(65 + j) # Level A, B, C...
    height = st.sidebar.number_input(f"Level {level_name} Net Height", min_value=1.0, value=float(group_data['bin_heights'][j]), key=f"level_{active_group_idx}_{j}")
    current_bin_heights.append(height)
group_data['bin_heights'] = current_bin_heights

st.sidebar.subheader("Materials & Appearance")
group_data['shelf_thickness'] = st.sidebar.number_input("Shelf Thickness (mm)", min_value=1.0, value=float(group_data['shelf_thickness']), key=f"shelf_thick_{active_group_idx}")
group_data['side_panel_thickness'] = st.sidebar.number_input("Outer Side Panel Thickness (mm)", min_value=1.0, value=float(group_data['side_panel_thickness']), key=f"side_panel_thick_{active_group_idx}")
group_data['color'] = st.sidebar.color_picker("Structure Color", value=group_data['color'], key=f"color_{active_group_idx}")
group_data['zoom'] = st.sidebar.slider("Zoom", 1.0, 5.0, group_data.get('zoom', 1.0), 0.1, key=f"zoom_{active_group_idx}", help="Increase to zoom out and see more area around the design.")

# --- Calculate and Display Final Height ---
total_net_bin_h = sum(group_data['bin_heights'])
num_shelves_for_calc = group_data['num_rows'] + (1 if group_data['has_top_cap'] else 0)
total_shelf_h = num_shelves_for_calc * group_data['shelf_thickness']
calculated_total_height = total_net_bin_h + total_shelf_h + group_data['ground_clearance']
st.sidebar.metric("Calculated Total Height", f"{calculated_total_height:.1f} mm")
group_data['total_height'] = calculated_total_height 

# --- Main Area for Drawing ---
st.header(f"Generated Design for: {group_data['name']}")
fig = draw_bay_group(group_data)
st.pyplot(fig, use_container_width=True)

# --- Global Download Button (FIXED) ---
st.sidebar.markdown("---")
st.sidebar.header("Download All Designs")

# Create a placeholder for the download button
download_button_placeholder = st.sidebar.empty()

if st.sidebar.button("Generate PPTX"):
    # Recalculate all group heights before creating the buffer
    for group in st.session_state.bay_groups:
        total_net_bin_h = sum(group['bin_heights'])
        num_shelves_for_calc = group['num_rows'] + (1 if group['has_top_cap'] else 0)
        total_shelf_h = num_shelves_for_calc * group['shelf_thickness']
        group['total_height'] = total_net_bin_h + total_shelf_h + group['ground_clearance']

    ppt_buffer = create_editable_powerpoint(st.session_state.bay_groups)
    
    download_button_placeholder.download_button(
        label="Download Now",
        data=ppt_buffer,
        file_name="all_bay_designs.pptx",
        mime="application/vnd.openxmlformats-officedocument.presentationml.presentation"
    )
