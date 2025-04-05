from PIL import Image
import os

def create_icon():
    # Create a simple colored square as icon
    sizes = [(16,16), (32,32), (48,48), (64,64), (128,128), (256,256)]
    
    # Create a blue gradient icon
    icon = Image.new('RGB', (256,256), color='white')
    pixels = icon.load()
    
    # Create a simple gradient
    for x in range(256):
        for y in range(256):
            pixels[x,y] = (0, int(x/2), int(y/2))  # Blue gradient
            
    # Save icon in different sizes
    icon_images = []
    for size in sizes:
        icon_images.append(icon.resize(size, Image.Resampling.LANCZOS))
    
    # Save as ICO file
    icon_images[0].save(
        'icon.ico',
        format='ICO',
        sizes=sizes,
        append_images=icon_images[1:]
    )

if __name__ == '__main__':
    create_icon()
    print("Icon created successfully!")
