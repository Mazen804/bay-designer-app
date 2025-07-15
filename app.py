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
st.markdown("Use the sidebar to manage bay groups. Configure each group in its tab below.")

# --- Helper Functions ---

def draw_dimension_line(ax, x1, y1, x2, y2, text, is_vertical=False, offset=10):
    """Draws a dimension line with arrows and text on the matplotlib axis."""
    ax.plot([x1, x2], [y1, y2], color='black', lw=1)
    if is_vertical:
        ax.plot(x1, y1, marker='v', color='black', markersize=5)
        ax.plot(x2, y2, marker='^', color='black', markersize=5)
        ax.text(x1 + offset, (y1 + y2) / 2, text, va='center', ha='left', fontsize=9, rotation=90)
    else:
        ax.plot(x1, y1, marker='<', color='black', markersize=5)
        ax.plot(x2, y2, marker='>', color='black', markersize=5)
        ax.text((x1 + x2) / 2, y1 + offset, text, va='bottom', ha='center', fontsize=9)

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

    # --- Calculations (MODIFIED) ---
    # Total width of the group, assuming dividers between bays have zero thickness.
    total_group_width = (num_bays * bay_width) + (2 * side_panel_thickness)
    
    fig, ax = plt.subplots(figsize=(12, 8))

    # --- Draw Structure (MODIFIED) ---
    structure_height = total_height - ground_clearance
    
    # Draw left-most side panel
    ax.add_patch(patches.Rectangle((0, 0), side_panel_thickness, total_height, facecolor=color))
    current_x = side_panel_thickness

    # Loop through each bay in the group
    for bay_idx in range(num_bays):
        # Calculate bin widths for this specific bay
        net_width_per_bay = bay_width
        total_internal_dividers = (num_cols - 1) * shelf_thickness
        bin_width = (net_width_per_bay - total_internal_dividers) / num_cols if num_cols > 0 else 0

        # Draw internal bin-splits
        bin_start_x = current_x
        if num_cols > 1:
            for i in range(1, num_cols):
                split_x = bin_start_x + (i * bin_width) + ((i-1) * shelf_thickness)
                ax.add_patch(patches.Rectangle((split_x, ground_clearance), shelf_thickness, structure_height, facecolor=color))
        
        # Add a thin visual line to separate bays if it's not the last one
        if bay_idx < num_bays - 1:
             divider_x = current_x + bay_width
             ax.plot([divider_x, divider_x], [ground_clearance, structure_height], color='#aaaaaa', lw=1, linestyle='--')

        current_x += bay_width

    # Draw the final right-most side panel
    ax.add_patch(patches.Rectangle((current_x, 0), side_panel_thickness, total_height, facecolor=color))

    # --- Draw Horizontal Shelves (span across the entire group) ---
    current_y = ground_clearance
    for i in range(num_rows):
        ax.add_patch(patches.Rectangle((0, current_y), total_group_width, shelf_thickness, facecolor=color))
        current_y += shelf_thickness
        
        if i < len(bin_heights):
            bin_h = bin_heights[num_rows - 1 - i]
            level_name = chr(65 + i)
            ax.text(-40, current_y + bin_h / 2, level_name, va='center', ha='center', fontsize=12, fontweight='bold')
            current_y += bin_h

    if has_top_cap:
        ax.add_patch(patches.Rectangle((0, total_height - shelf_thickness), total_group_width, shelf_thickness, facecolor=color))

    # --- Draw Dimension Lines ---
    dim_offset = 0.05 * max(total_height, total_group_width)
    draw_dimension_line(ax, 0, -dim_offset, total_group_width, -dim_offset, f"Total Group Width: {total_group_width:.0f} mm")
    draw_dimension_line(ax, -dim_offset, 0, -dim_offset, total_height, f"Total Height: {total_height:.0f} mm", is_vertical=True)

    # --- Final Touches ---
    ax.set_aspect('equal', adjustable='box')
    ax.axis('off')
    ax.set_xlim(-dim_offset * 2, total_group_width + dim_offset * 2)
    ax.set_ylim(-dim_offset * 2, total_height + dim_offset * 2)
    
    return fig

def create_powerpoint(figures_with_names):
    """Creates a PowerPoint presentation from a list of figures."""
    prs = Presentation()
    for fig, name in figures_with_names:
        slide_layout = prs.slide_layouts[5]  # Blank slide layout
        slide = prs.slides.add_slide(slide_layout)
        
        title = slide.shapes.title
        title.text = f"Design for: {name}"

        buf = io.BytesIO()
        fig.savefig(buf, format='png', bbox_inches='tight')
        buf.seek(0)
        
        # Add picture to slide, centered
        pic_width_in = 7.5
        pic_height_in = pic_width_in * (fig.get_figheight() / fig.get_figwidth())
        left = Inches((prs.slide_width.inches - pic_width_in) / 2)
        top = Inches((prs.slide_height.inches - pic_height_in) / 2)
        
        slide.shapes.add_picture(buf, left, top, width=Inches(pic_width_in))

    # Save presentation to a buffer
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
        "bin_heights": [350.0] * 5
    }]

# --- Sidebar for Managing Groups ---
with st.sidebar:
    st.header("Manage Bay Groups")
    
    with st.form("new_group_form"):
        new_group_name = st.text_input("New Group Name", "New Group")
        add_group_submitted = st.form_submit_button("Add Group")
        if add_group_submitted:
            new_group = st.session_state.bay_groups[0].copy() # Copy from the first group as a template
            new_group['name'] = new_group_name
            st.session_state.bay_groups.append(new_group)
            st.rerun()

    if len(st.session_state.bay_groups) > 1:
        if st.button("Remove Last Group"):
            st.session_state.bay_groups.pop()
            st.rerun()

# --- Main Area with Tabs for Each Group ---
if not st.session_state.bay_groups:
    st.warning("No bay groups to display. Please add one using the sidebar.")
else:
    group_names = [g['name'] for g in st.session_state.bay_groups]
    tabs = st.tabs(group_names)
    
    all_figures = []

    for i, tab in enumerate(tabs):
        with tab:
            group_data = st.session_state.bay_groups[i]
            st.header(f"Configuration for: {group_data['name']}")
            
            # --- Configuration Controls within each tab ---
            c1, c2, c3 = st.columns(3)
            with c1:
                st.subheader("Structure")
                group_data['num_bays'] = st.number_input("Number of Bays in Group", min_value=1, value=group_data['num_bays'], key=f"num_bays_{i}")
                group_data['bay_width'] = st.number_input("Width per Bay (mm)", min_value=1, value=group_data['bay_width'], key=f"bay_width_{i}")
                group_data['total_height'] = st.number_input("Total Height (mm)", min_value=1, value=group_data['total_height'], key=f"total_height_{i}")
                group_data['ground_clearance'] = st.number_input("Ground Clearance (mm)", min_value=0, value=group_data['ground_clearance'], key=f"ground_clearance_{i}")
                group_data['has_top_cap'] = st.checkbox("Add Top Cap", value=group_data['has_top_cap'], key=f"has_top_cap_{i}")

            with c2:
                st.subheader("Layout")
                group_data['num_rows'] = st.number_input("Shelves (Rows)", min_value=1, value=group_data['num_rows'], key=f"num_rows_{i}")
                group_data['num_cols'] = st.number_input("Bin-Split (Columns)", min_value=1, value=group_data['num_cols'], key=f"num_cols_{i}")
                
                # Dynamic bin height inputs
                st.markdown("**Individual Shelf Heights**")
                new_bin_heights = []
                # Ensure bin_heights list is the correct length
                while len(group_data['bin_heights']) < group_data['num_rows']:
                    group_data['bin_heights'].append(350.0)
                while len(group_data['bin_heights']) > group_data['num_rows']:
                    group_data['bin_heights'].pop()

                for j in range(group_data['num_rows']):
                    level_name = chr(65 + (group_data['num_rows'] - 1 - j))
                    height = st.number_input(f"Level {level_name} Height", min_value=1.0, value=group_data['bin_heights'][j], key=f"level_{i}_{j}")
                    new_bin_heights.append(height)
                group_data['bin_heights'] = new_bin_heights

            with c3:
                st.subheader("Materials & Appearance")
                group_data['shelf_thickness'] = st.number_input("Shelf & Bin-Split Thickness (mm)", min_value=1, value=group_data['shelf_thickness'], key=f"shelf_thick_{i}")
                group_data['side_panel_thickness'] = st.number_input("Outer Side Panel Thickness (mm)", min_value=1, value=group_data['side_panel_thickness'], key=f"side_panel_thick_{i}")
                group_data['color'] = st.color_picker("Structure Color", value=group_data['color'], key=f"color_{i}")

            # --- Drawing Area ---
            st.markdown("---")
            st.subheader("Generated Design")
            
            fig = draw_bay_group(group_data)
            all_figures.append((fig, group_data['name']))
            st.pyplot(fig, use_container_width=True)

    # --- Global Download Button ---
    st.sidebar.markdown("---")
    st.sidebar.header("Download All Designs")
    if all_figures:
        ppt_buffer = create_powerpoint(all_figures)
        st.sidebar.download_button(
            label="Download All Designs (PPTX)",
            data=ppt_buffer,
            file_name="all_bay_designs.pptx",
            mime="application/vnd.openxmlformats-officedocument.presentationml.presentation"
        )
