import streamlit as st
import matplotlib.pyplot as plt
import matplotlib.patches as patches
import io
from pptx import Presentation
from pptx.util import Inches

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
    """Main function to draw a group of bays using Matplotlib."""
    # Unpack parameters
    num_bays = params['num_bays']
    bay_width = params['bay_width']
    total_height = params['total_height']
    ground_clearance = params['ground_clearance']
    shelf_thickness = params['shelf_thickness']
    side_panel_thickness = params['side_panel_thickness']
    num_cols = params['num_cols'] # This is the bin-split
    num_rows = params['num_rows'] # This is the shelves
    has_top_cap = params['has_top_cap']
    color = params['color']
    bin_heights = params['bin_heights']
    zoom_factor = params.get('zoom', 1.0)

    # --- Calculations ---
    total_group_width = (num_bays * bay_width) + (2 * side_panel_thickness)
    
    fig, ax = plt.subplots(figsize=(12, 12))

    # --- Draw Structure ---
    structure_height = total_height - ground_clearance
    
    # Draw left-most side panel
    ax.add_patch(patches.Rectangle((0, 0), side_panel_thickness, total_height, facecolor=color))
    current_x = side_panel_thickness

    # Loop through each bay in the group
    for bay_idx in range(num_bays):
        net_width_per_bay = bay_width
        total_internal_dividers = (num_cols - 1) * shelf_thickness
        bin_width = (net_width_per_bay - total_internal_dividers) / num_cols if num_cols > 0 else 0

        bin_start_x = current_x
        if num_cols > 1:
            for i in range(1, num_cols):
                split_x = bin_start_x + (i * bin_width) + ((i-1) * shelf_thickness)
                ax.add_patch(patches.Rectangle((split_x, ground_clearance), shelf_thickness, structure_height, facecolor=color))
        
        if bay_idx < num_bays - 1:
             divider_x = current_x + bay_width
             ax.plot([divider_x, divider_x], [ground_clearance, structure_height], color='#aaaaaa', lw=1, linestyle='--')

        current_x += bay_width

    # Draw the final right-most side panel
    ax.add_patch(patches.Rectangle((current_x, 0), side_panel_thickness, total_height, facecolor=color))

    # --- Draw Horizontal Shelves & Bin Height Dimensions ---
    current_y = ground_clearance
    dim_offset_x = 0.05 * total_group_width

    for i in range(num_rows):
        ax.add_patch(patches.Rectangle((0, current_y), total_group_width, shelf_thickness, facecolor=color))
        shelf_top_y = current_y + shelf_thickness
        
        if i < len(bin_heights):
            bin_h = bin_heights[num_rows - 1 - i]
            level_name = chr(65 + i)
            
            # Draw bin height dimension line
            bin_bottom_y = shelf_top_y
            bin_top_y = bin_bottom_y + bin_h
            draw_dimension_line(ax, total_group_width + dim_offset_x, bin_bottom_y, total_group_width + dim_offset_x, bin_top_y, f"{bin_h:.0f}", is_vertical=True, offset=5, color='#3b82f6')
            
            # Draw Level Name
            ax.text(-dim_offset_x, (bin_bottom_y + bin_top_y) / 2, level_name, va='center', ha='center', fontsize=12, fontweight='bold')
            
            current_y = bin_top_y

    if has_top_cap:
        ax.add_patch(patches.Rectangle((0, total_height - shelf_thickness), total_group_width, shelf_thickness, facecolor=color))

    # --- Draw Main Dimension Lines ---
    dim_offset_y = 0.05 * total_height
    # Total Width
    draw_dimension_line(ax, 0, -dim_offset_y, total_group_width, -dim_offset_y, f"Total Group Width: {total_group_width:.0f} mm", offset=10)
    # Total Height (moved further out)
    draw_dimension_line(ax, -dim_offset_x * 2.5, 0, -dim_offset_x * 2.5, total_height, f"Total Height: {total_height:.0f} mm", is_vertical=True, offset=10)

    # --- Final Touches ---
    ax.set_aspect('equal', adjustable='box')
    ax.axis('off')
    # Use zoom factor to adjust the view
    ax.set_xlim(-dim_offset_x * 4 * zoom_factor, total_group_width + dim_offset_x * 4 * zoom_factor)
    ax.set_ylim(-dim_offset_y * 2 * zoom_factor, total_height + dim_offset_y * 2 * zoom_factor)
    
    return fig

def create_powerpoint(figures_with_names):
    """Creates a PowerPoint presentation from a list of figures."""
    prs = Presentation()
    for fig, name in figures_with_names:
        blank_slide_layout = prs.slide_layouts[6]
        slide = prs.slides.add_slide(blank_slide_layout)
        
        title_shape = slide.shapes.add_textbox(Inches(0.5), Inches(0.2), Inches(9), Inches(0.5))
        title_shape.text = f"Design for: {name}"

        buf = io.BytesIO()
        fig.savefig(buf, format='png', bbox_inches='tight', pad_inches=0.2)
        buf.seek(0)
        
        margin = Inches(0.25)
        max_width = prs.slide_width - (2 * margin)
        max_height = prs.slide_height - Inches(0.75) - margin 

        img_width_px, img_height_px = fig.get_size_inches() * fig.dpi
        img_aspect_ratio = img_width_px / img_height_px

        if (max_width / img_aspect_ratio) > max_height:
            pic_height = max_height
            pic_width = max_height * img_aspect_ratio
        else:
            pic_width = max_width
            pic_height = max_width / img_aspect_ratio

        left = (prs.slide_width - pic_width) / 2
        top = Inches(0.75) + ((max_height - pic_height) / 2)
        
        slide.shapes.add_picture(buf, left, top, width=pic_width, height=pic_height)

    ppt_buf = io.BytesIO()
    prs.save(ppt_buf)
    ppt_buf.seek(0)
    return ppt_buf

# --- Initialize Session State ---
if 'bay_groups' not in st.session_state:
    st.session_state.bay_groups = [{
        "name": "Group A", "num_bays": 2, "bay_width": 1050, "total_height": 2000,
        "ground_clearance": 50, "shelf_thickness": 18, "side_panel_thickness": 18,
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

# Select which group to edit
group_names = [g['name'] for g in st.session_state.bay_groups]
selected_group_name = st.sidebar.selectbox("Select Group to Edit", group_names)
active_group_idx = group_names.index(selected_group_name)
group_data = st.session_state.bay_groups[active_group_idx]

# --- Configuration for the selected group ---
st.sidebar.header(f"Configuration for: {group_data['name']}")

st.sidebar.subheader("Structure")
group_data['num_bays'] = st.sidebar.number_input("Number of Bays in Group", min_value=1, value=group_data['num_bays'], key=f"num_bays_{active_group_idx}")
group_data['bay_width'] = st.sidebar.number_input("Width per Bay (mm)", min_value=1, value=group_data['bay_width'], key=f"bay_width_{active_group_idx}")
group_data['total_height'] = st.sidebar.number_input("Total Height (mm)", min_value=1, value=group_data['total_height'], key=f"total_height_{active_group_idx}")
group_data['ground_clearance'] = st.sidebar.number_input("Ground Clearance (mm)", min_value=0, value=group_data['ground_clearance'], key=f"ground_clearance_{active_group_idx}")
group_data['has_top_cap'] = st.sidebar.checkbox("Add Top Cap", value=group_data['has_top_cap'], key=f"has_top_cap_{active_group_idx}")

st.sidebar.subheader("Layout")
group_data['num_rows'] = st.sidebar.number_input("Shelves (Rows)", min_value=1, value=group_data['num_rows'], key=f"num_rows_{active_group_idx}")
group_data['num_cols'] = st.sidebar.number_input("Bin-Split (Columns)", min_value=1, value=group_data['num_cols'], key=f"num_cols_{active_group_idx}")

st.sidebar.markdown("**Individual Shelf Heights**")
new_bin_heights = []
while len(group_data['bin_heights']) < group_data['num_rows']:
    group_data['bin_heights'].append(350.0)
while len(group_data['bin_heights']) > group_data['num_rows']:
    group_data['bin_heights'].pop()

for j in range(group_data['num_rows']):
    level_name = chr(65 + (group_data['num_rows'] - 1 - j))
    height = st.sidebar.number_input(f"Level {level_name} Height", min_value=1.0, value=group_data['bin_heights'][j], key=f"level_{active_group_idx}_{j}")
    new_bin_heights.append(height)
group_data['bin_heights'] = new_bin_heights

st.sidebar.subheader("Materials & Appearance")
group_data['shelf_thickness'] = st.sidebar.number_input("Shelf Thickness (mm)", min_value=1, value=group_data['shelf_thickness'], key=f"shelf_thick_{active_group_idx}")
group_data['side_panel_thickness'] = st.sidebar.number_input("Outer Side Panel Thickness (mm)", min_value=1, value=group_data['side_panel_thickness'], key=f"side_panel_thick_{active_group_idx}")
group_data['color'] = st.sidebar.color_picker("Structure Color", value=group_data['color'], key=f"color_{active_group_idx}")
group_data['zoom'] = st.sidebar.slider("Zoom", 1.0, 5.0, group_data.get('zoom', 1.0), 0.1, key=f"zoom_{active_group_idx}", help="Increase to zoom out and see more area around the design.")


# --- Main Area for Drawing ---
st.header(f"Generated Design for: {group_data['name']}")
fig = draw_bay_group(group_data)
st.pyplot(fig, use_container_width=True)

# --- Global Download Button ---
st.sidebar.markdown("---")
st.sidebar.header("Download All Designs")

# We need to generate all figures for the download button, not just the active one
all_figures = []
for group in st.session_state.bay_groups:
    fig_to_download = draw_bay_group(group)
    all_figures.append((fig_to_download, group['name']))
    plt.close(fig_to_download) # Close the figure to save memory

if all_figures:
    ppt_buffer = create_powerpoint(all_figures)
    st.sidebar.download_button(
        label="Download All Designs (PPTX)",
        data=ppt_buffer,
        file_name="all_bay_designs.pptx",
        mime="application/vnd.openxmlformats-officedocument.presentationml.presentation"
    )
