import streamlit as st
from streamlit_drawable_canvas import st_canvas
from pptx.util import Inches
from PIL import Image
import pandas as pd
from pptx import Presentation
from io import BytesIO
import pptx_replace as pr
prs = Presentation('SIQ Methodology 1.pptx')
slide= prs.slides[0]

def update_shape_text(shape, value):
    """
    Preserve formatting where possible.
    """

    p = shape.text_frame.paragraphs[0]

    # Best case: existing run with formatting
    if len(p.runs) > 0:
        p.runs[0].text = str(value)

        # Remove any additional runs that might contain old text
        for run in p.runs[1:]:
            run.text = ""

    # Fallback: no runs exist
    else:
        p.text = str(value)
clutter_multiplier_dict ={
'1':1.05,
'2':1,
'3':0.95,
'4':0.9,
'5':0.85,
'6':0.8,
'7':0.8,
'8' :0.8,
'9'	:0.8,
'10':0.8
}

# Can potentially make this a drop down that when it is selected it takes the key and uses that to Insert the Value
asset_category_dict = {
"Team/Athlete Clothing": 1.25,
"Team Vehicle":1.20,
"Award"	:1.15,
"On-Pitch/Course":1.125,
"Graphics":1.10,
"In-Stadia"	:1.075,
"Other Venue":1.05,
"Equipment/Uniform":1.04,
"Other":	1.00
}

st.markdown("## Image Boudning Box & SIQ Slide Measurement Tool", text_alignment='center')

st.markdown('### Instrcutions')
st.write(
    "1. Ensure image is either a .png, .jpeg or .png file.\n"
    "2. Input the number of logos for the brand you are tracking.\n"
    "3. Draw boxes over the tracked logos then draw boxes over all other logos visible in the picture.\n"
    "4. Then select the Asset Category that best fits the logo.\n"
    "5. Download and inspect slide."
)
#Upload Image to be measured and draw boxes which allow unit conversion pix -> cm 0.0265
uploaded_file = st.file_uploader("Upload an Image", type=["png", "jpg", "jpeg"])

scale_unit = 0.0265
num_logo = st.number_input('Number of instances of tracked Logos:', value=0, max_value=100,format="%.0f")

height_list = []
width_list = []

if uploaded_file is not None:
    image = Image.open(uploaded_file)
    img_width, img_height = image.size

    st.write("Draw a rectangle around the object to measure:")

    canvas_result = st_canvas(
        fill_color="rgba(255,165,0,0.3)",
        stroke_width=2,
        stroke_color="#FF4B4B",
        background_image=image,
        update_streamlit=True,
        height=img_height,
        width=img_width,
        drawing_mode="rect",
        key='canvas'
    )
    if canvas_result.json_data is not None:
        objects = canvas_result.json_data["objects"]
        if objects:
            for box in objects:
                box_width_px = box["width"] * box["scaleX"]
                box_height_px = box["height"] * box["scaleY"]

                real_width = box_width_px * scale_unit
                real_height = box_height_px * scale_unit

                height_list.append(real_height)
                width_list.append(real_width)

# Store measurements in the df and calc the logo size perc
image_df = pd.DataFrame(columns=['Height', 'Width'])
image_df['Height'] = height_list
image_df['Width'] = width_list

image_df['Logo Size'] = image_df['Height'] * image_df['Width']
#st.dataframe(image_df)
if uploaded_file is not None:
    image_size = (img_height*scale_unit) * (img_width*scale_unit)

    logo_size_perc = (image_df.iloc[0,2] /image_size *100)

    BASE = 2

    asset_category = st.selectbox(label='Select Asset Category from dropdown', options=asset_category_dict.keys())

    base_multiplier = BASE * asset_category_dict[asset_category]
    size_factor = logo_size_perc * 0.03
    num_logo_factor = num_logo *0.15
    clutter_factor = clutter_multiplier_dict[str(len(image_df))]

    zero_five_Score = round((base_multiplier + size_factor + num_logo_factor) * clutter_factor,2)

    siq = 0.5 + 1.3 * (zero_five_Score)

    shape_names = ['SIQ Score', 'SIQ Score / 10', 'Formula','df length', 'num logo', 'screen perc', 'screen perc' ]
    for shape in slide.shapes:
        if shape.name == 'SIQ Discount':
            siq_discount = siq*10
            update_shape_text(shape,f"{siq_discount:.2f}% of AVE retained")
        elif shape.name == 'SIQ Score':
            update_shape_text(shape,f"FINAL SIQ+ {siq:.3f}")
        elif shape.name == 'SIQ Score / 10':
            update_shape_text(shape, f"AUTOMATED SIQ {siq:.3f}/10")
        elif shape.name == 'Formula':
            update_shape_text(shape, f"({BASE} + (0.03 x {logo_size_perc:.2f}) + (0.15 x {num_logo})) x {round(clutter_factor * 100, 0)}%")
        elif shape.name == 'df length':
            update_shape_text(shape, len(image_df))
        elif shape.name == 'num logo':
            update_shape_text(shape, num_logo)
        elif shape.name == 'screen perc':
            update_shape_text(shape, f"{logo_size_perc:.2f}%")  
        elif shape.name == 'Image 1': 
            pr.replace_picture(prs.slides[0], uploaded_file, auto_reshape=True)


    buffer = BytesIO()
    prs.save(buffer)
    buffer.seek(0)

    st.download_button(
        label="Download Powerpoint",
        data= buffer,
        file_name="presentation.pptx",
        mime="application/vnd.openxmlformats-officedocument.presentationml.presentation"
    )

