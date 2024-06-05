from ppt_image_inserter import PPTImageInserter
import matplotlib.pyplot as plt
import os

# Create a directory to save images
os.makedirs("images", exist_ok=True)


def plot_function(i, color):
    plt.figure()
    plt.plot([0, 1, 2, 3], [0, 1, 2, 3], c=color)
    plt.title(f"Plot {i+1}", fontsize=18)
    plt.xlabel("X-axis", fontsize=18)
    plt.ylabel("Y-axis", fontsize=18)


'''
Example 1
Save all images in a folder and then insert into PowerPoint. 
'''
# Generate and save some example plots
for i in range(12):
    plot_function(i, color='C0')
    image_path = f"images/plotA_{i+1}.png"
    plt.savefig(image_path)
    plt.close()

imageA_paths = [f"images/plotA_{i+1}.png" for i in range(12)]

# Create an instance of PPTImageInserter
ppt_inserter = PPTImageInserter(grid_dims=(
    3, 3), spacing=(0.05, 0.05), title_font_size=16)

# Add images
for image_path in imageA_paths:
    ppt_inserter.add_image(image_path, slide_title="Plot A")

ppt_inserter.save("exampleA.pptx")

'''
Example 2
Add image to the PowerPoint for each iteration of the image generation loop.
Use a 2x2 grid with no title. 
'''
ppt_inserter = PPTImageInserter(grid_dims=(
    2, 2), spacing=(0.05, 0.05), title_font_size=16)
for i in range(12):
    plot_function(i, color='C0')
    image_path = f"images/image.png"
    plt.savefig(image_path)
    ppt_inserter.add_image(image_path)
    plt.close()

ppt_inserter.save("exampleB.pptx")


'''
Example 3
Add new slides for different image set. 
'''
# Generate and save some example plots for a different set
for i in range(12):
    plot_function(i, color='C1')
    plt_path = f"images/plotB_{i+1}.png"
    plt.savefig(plt_path)
    plt.close()

# Define the paths to the images
imageA_paths = [f"images/plotA_{i+1}.png" for i in range(12)]
imageB_paths = [f"images/plotB_{i+1}.png" for i in range(12)]

ppt_inserter = PPTImageInserter(grid_dims=(
    3, 3), spacing=(0.05, 0.05), title_font_size=16)

# Add a new slide and insert images for Plot A
for image_path in imageA_paths:
    ppt_inserter.add_image(image_path, slide_title="Plot A")

# Add a new slide and insert images for Plot B
ppt_inserter.add_slide("Plot B")
for image_path in imageB_paths:
    ppt_inserter.add_image(image_path, slide_title="Plot B")

ppt_inserter.save("exampleC.pptx")
