import streamlit as st
import matplotlib.pyplot as plt
import matplotlib.patches as patches

# --- Page Configuration ---
st.set_page_config(layout="wide", page_title="Storage Bay Designer")

# --- Main Title ---
st.title("Storage Bay Designer")
st.markdown("Use the sidebar to enter your specifications and generate a bay design.")

# --- Helper Functions ---

def draw_dimension_line(ax, x1, y1, x2, y2, text, is_vertical=False, offset=10):
    """Draws a dimension line with arrows and text on the matplotlib axis."""
    ax.plot([x1, x2], [y1, y2], color='black', lw=1)
    
    # Draw arrows
    if is_vertical:
        ax.plot(x1, y1, marker='v', color='black', markersize=5)
        ax.plot(x2, y2, marker='^', color='black', markersize=5)
        ax.text(x1 + offset, (y1 + y2) / 2, text, va='center', ha='left', fontsize=9, rotation=90)
    else:
        ax.plot(x1, y1, marker='<', color='black', markersize=5)
        ax.plot(x2, y2, marker='>', color='black', markersize=5)
        ax.text((x1 + x2) / 2, y1 + offset, text, va='bottom', ha='center', fontsize=9)

def draw_bay(params):
    """Main function to draw the bay using Matplotlib."""
    fig, ax = plt.subplots(figsize=(10, 10))
    
    # Unpack parameters
    gross_width = params['gross_width']
    total_height = params['total_height']
    ground_clearance = params['ground_clearance']
    shelf_thickness = params['shelf_thickness']
    side_panel_thickness = params['side_panel_thickness']
    num_cols = params['num_cols']
    num_rows = params['num_rows']
    has_top_cap = params['has_top_cap']
    color = params['color']
    bin_heights = params['bin_heights']

    # --- Calculations ---
    net_width = gross_width - (2 * side_panel_thickness)
    total_internal_dividers_width = (num_cols - 1) * shelf_thickness if num_cols > 1 else 0
    available_width_for_bins = net_width - total_internal_dividers_width
    bin_width = available_width_for_bins / num_cols if num_cols > 0 else 0

    # --- Draw Structure ---
    structure_height = total_height - ground_clearance

    # Side Panels (as feet)
    ax.add_patch(patches.Rectangle((0, 0), side_panel_thickness, total_height, facecolor=color))
    ax.add_patch(patches.Rectangle((gross_width - side_panel_thickness, 0), side_panel_thickness, total_height, facecolor=color))

    # Internal Dividers
    if num_cols > 1:
        for i in range(1, num_cols):
            x_pos = side_panel_thickness + (i * bin_width) + ((i - 1) * shelf_thickness)
            ax.add_patch(patches.Rectangle((x_pos, ground_clearance), shelf_thickness, structure_height, facecolor=color))

    # Shelves and Level Names (drawn from bottom up)
    current_y = ground_clearance
    for i in range(num_rows):
        # Bottom shelf of the current bin
        ax.add_patch(patches.Rectangle((0, current_y), gross_width, shelf_thickness, facecolor=color))
        current_y += shelf_thickness
        
        # Check if bin_heights has enough elements
        if i < len(bin_heights):
            bin_h = bin_heights[num_rows - 1 - i]
            
            # Add Level Name
            level_name = chr(65 + i)
            ax.text(-40, current_y + bin_h / 2, level_name, va='center', ha='center', fontsize=12, fontweight='bold')
            
            current_y += bin_h

    # Top Cap
    if has_top_cap:
        ax.add_patch(patches.Rectangle((0, total_height - shelf_thickness), gross_width, shelf_thickness, facecolor=color))

    # --- Draw Dimension Lines ---
    dim_offset = 0.05 * max(total_height, gross_width) # Make offset relative to max dimension
    draw_dimension_line(ax, 0, -dim_offset, gross_width, -dim_offset, f"{gross_width:.0f} mm")
    draw_dimension_line(ax, -dim_offset, 0, -dim_offset, total_height, f"{total_height:.0f} mm", is_vertical=True)
    
    if ground_clearance > 0:
        draw_dimension_line(ax, gross_width + dim_offset, 0, gross_width + dim_offset, ground_clearance, f"{ground_clearance:.0f} mm", is_vertical=True)

    if bin_width > 0:
        bin_start_x = side_panel_thickness
        bin_end_x = bin_start_x + bin_width
        draw_dimension_line(ax, bin_start_x, total_height + dim_offset, bin_end_x, total_height + dim_offset, f"{bin_width:.1f} mm")


    # --- Final Touches ---
    ax.set_aspect('equal', adjustable='box')
    ax.axis('off')
    ax.set_xlim(-dim_offset * 2, gross_width + dim_offset * 2)
    ax.set_ylim(-dim_offset * 2, total_height + dim_offset * 2)
    
    return fig, net_width, bin_width

# --- Sidebar Controls ---
with st.sidebar:
    st.header("Parameters")

    # Bay Dimensions Section
    st.subheader("Bay Dimensions")
    gross_width = st.number_input("Gross Width (mm)", min_value=1, value=1050)
    
    # Material Thickness Section
    st.subheader("Material Thickness")
    shelf_thickness = st.number_input("Shelves (mm)", min_value=1, value=18)
    side_panel_thickness = st.number_input("Side Panels (mm)", min_value=1, value=18)

    # Configuration Section
    st.subheader("Configuration")
    num_cols = st.number_input("Columns", min_value=1, value=4)
    num_rows = st.number_input("Rows", min_value=1, value=5, key="num_rows")
    has_top_cap = st.checkbox("Add Top Cap", value=True)
    ground_clearance = st.number_input("Ground Clearance (mm)", min_value=0, value=50)

    # --- Logic for managing heights ---
    
    # This function will be called when Total Height is changed
    def distribute_total_height():
        num_shelves = st.session_state.num_rows + (1 if has_top_cap else 0)
        total_shelf_h = num_shelves * shelf_thickness
        available_space = st.session_state.total_height_input - ground_clearance - total_shelf_h
        if available_space > 0:
            uniform_h = available_space / st.session_state.num_rows
            # Update the session state for each level height
            for i in range(st.session_state.num_rows):
                st.session_state[f"level_{i}"] = uniform_h

    # This function calculates total height from individual levels
    def get_calculated_total_height():
        total_bin_h = 0
        for i in range(num_rows):
            total_bin_h += st.session_state.get(f"level_{i}", 0)
        num_shelves_calc = num_rows + (1 if has_top_cap else 0)
        return total_bin_h + (num_shelves_calc * shelf_thickness) + ground_clearance
    
    # Initialize level heights if they don't exist
    for i in range(num_rows):
        if f"level_{i}" not in st.session_state:
            st.session_state[f"level_{i}"] = 350.0

    # Total Height input now calls the distribution function on change
    total_height = st.number_input(
        "Total Height (mm)", 
        min_value=1, 
        value=int(get_calculated_total_height()), 
        key="total_height_input",
        on_change=distribute_total_height
    )

    # --- Dynamic Level Height Inputs ---
    st.subheader("Level Heights")
    
    # Create input fields for each level
    current_bin_heights = []
    for i in range(num_rows):
        level_name = chr(65 + (num_rows - 1 - i))
        level_height = st.number_input(
            f"Level {level_name} Height (mm)", 
            min_value=1.0, 
            value=float(st.session_state.get(f"level_{i}", 350.0)),
            key=f"level_{i}"
        )
        current_bin_heights.append(level_height)

    # Appearance Section
    st.subheader("Appearance")
    color = st.color_picker("Structure Color", "#4A90E2")


# --- Main Area ---
final_total_height = get_calculated_total_height()

col1, col2 = st.columns([1, 2.5])

with col1:
    st.subheader("Calculated Dimensions")
    
    # Pack parameters for drawing function
    params = {
        "gross_width": gross_width,
        "total_height": final_total_height,
        "ground_clearance": ground_clearance,
        "shelf_thickness": shelf_thickness,
        "side_panel_thickness": side_panel_thickness,
        "num_cols": num_cols,
        "num_rows": num_rows,
        "has_top_cap": has_top_cap,
        "color": color,
        "bin_heights": current_bin_heights
    }

    # Only draw if the height is positive
    if final_total_height > 0 and gross_width > 0:
        fig, net_width, bin_width = draw_bay(params)
        
        st.metric("Net Width", f"{net_width:.1f} mm")
        st.metric("Calculated Bin Width", f"{bin_width:.1f} mm")
        st.metric("Final Total Height", f"{final_total_height:.1f} mm")
    else:
        st.error("Invalid dimensions. Please check your inputs.")
        fig = None


with col2:
    if fig:
        st.pyplot(fig)

