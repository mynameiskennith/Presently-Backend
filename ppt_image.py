import os
import tempfile
from io import BytesIO
from PIL import Image
import aspose.slides as slides
import aspose.pydrawing as drawing
import base64

def convert_ppt_to_images(ppt_file_contents):
    """
    Convert PPTX slides to an array of Base64-encoded strings for FastAPI responses.

    Args:
        ppt_file_contents (bytes): Content of the uploaded PPTX file.

    Returns:
        dict: A dictionary containing a list of Base64-encoded image strings.
    """
    # Save the uploaded file to a temporary file
    with tempfile.NamedTemporaryFile(delete=False, suffix=".pptx") as tmp_ppt_file:
        tmp_ppt_file.write(ppt_file_contents)
        tmp_ppt_file_path = tmp_ppt_file.name

    try:
        # Load the presentation
        pres = slides.Presentation(tmp_ppt_file_path)

        # List to store Base64-encoded image strings
        image_blobs_base64 = []

        # Create images and encode them in Base64
        with tempfile.TemporaryDirectory() as tmpdirname:
            for slide in pres.slides:
                # Generate a thumbnail for the slide
                bmp = slide.get_thumbnail(1, 1)

                # Save the thumbnail to a temporary file
                filename = f"Slide_{slide.slide_number}.jpg"
                filepath = os.path.join(tmpdirname, filename)
                bmp.save(filepath, drawing.imaging.ImageFormat.jpeg)

                # Convert the image to Base64
                with BytesIO() as img_stream:
                    with Image.open(filepath) as img:
                        img.save(img_stream, format='PNG')  # Convert to PNG format
                        image_blob = img_stream.getvalue()  # Get the binary content
                        image_base64 = base64.b64encode(image_blob).decode('utf-8')  # Base64 encode and decode to string
                        image_blobs_base64.append(image_base64)  # Append to the list

        return {"images": image_blobs_base64}

    finally:
        # Clean up the temporary file
        if os.path.exists(tmp_ppt_file_path):
            os.remove(tmp_ppt_file_path)

# import os
# import tempfile
# from io import BytesIO
# from PIL import Image
# # from spire.presentation import Presentation
# # from spire.presentation import *
# # from spire.presentation.common import *
# from PIL import Image
# import aspose.slides as slides
# import aspose.pydrawing as drawing

# def convert_ppt_to_images(ppt_file_contents):
#     # Save the uploaded file to a temporary file
#     with tempfile.NamedTemporaryFile(delete=False, suffix=".pptx") as tmp_ppt_file:
#         tmp_ppt_file.write(ppt_file_contents)  # Write the content of the uploaded file
#         tmp_ppt_file_path = tmp_ppt_file.name    # Get the file path

#     try:
        

#         pres = slides.Presentation(tmp_ppt_file_path)

#         images = []
#         with tempfile.TemporaryDirectory() as tmpdirname:
#             for sld in pres.slides:
#                 bmp = sld.get_thumbnail(1, 1)
#                 filename = f"Slide_{sld.slide_number}.jpg"
#                 filepath = f"{tmpdirname}/{filename}"
#                 bmp.save(f"{filepath}", drawing.imaging.ImageFormat.jpeg)

#                 imgStream = BytesIO()
                
#                 with Image.open(filepath) as img:
#                     img.save(imgStream,format='PNG')
#                     images.append(imgStream)
#                 print("////////////////////////////////////////////##########################################")
#                 print(images)
#                 return {"images":images}





#         # ppt_stream = BytesIO()
#         # prs.save(ppt_stream)
#         # ppt_stream.seek(0)  # Reset stream position for reading
#         # return ppt_stream

#         # Load the presentation from the saved file
#         # presentation = Presentation()
#         # presentation.LoadFromFile(tmp_ppt_file_path)

#         # images = []
#         # with tempfile.TemporaryDirectory() as tmpdirname:
#             # presentation.SaveToFile("PresentationToPDF.pdf", FileFormat.PDF)
#             # for i, slide in enumerate(presentation.Slides):
#             #     fileName = f"slide_{i}.png"
#             #     con(slide,fileName)
#                 # image = slide.SaveAsImage()
#                 # # file_path = os.path.join(tmpdirname, fileName)
#                 # image.Save(fileName)
#                 # image.Dispose()

#                 # Open the image, convert it to RGB (to ensure compatibility), and store it in memory
#                 # with Image.open(fileName) as img:
#                 #     print("////////////////////////////////////////////////////////////////////////////")
#                 #     print(type(img))
#                     # images.append(img.copy().convert('RGB'))

#         # presentation.Dispose()
#         # return {'images':images}

#     finally:
#         # Clean up the temporary file
#         if os.path.exists(tmp_ppt_file_path):
#             os.remove(tmp_ppt_file_path)