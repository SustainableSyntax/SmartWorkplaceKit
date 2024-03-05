from PIL import Image, ImageDraw, ImageFont

def create_banner_with_gradient(image_path, output_path, gradient_color1, gradient_color2, banner_height): 
    # Öffnen Sie das Bild 
    logo = Image.open(image_path)
    # Erstellen Sie eine neue leere Bilddatei mit der Größe des Banners
    banner_width = logo.width banner = Image.new('RGB', (banner_width, banner_height), color=gradient_color1)
    # Zeichnen Sie den farbigen Gradienten auf das Banner
    draw = ImageDraw.Draw(banner) for y in range(banner_height): draw.line([(0, y), (banner_width, y)], fill=( int(gradient_color1[0] + (gradient_color2[0] - gradient_color1[0]) * y / banner_height), int(gradient_color1[1] + (gradient_color2[1] - gradient_color1[1]) * y / banner_height), int(gradient_color1[2] + (gradient_color2[2] - gradient_color1[2]) * y / banner_height) ))
    # Platzieren Sie das Logo in der Mitte des Banners
    logo_position = ((banner_width - logo.width) // 2, (banner_height - logo.height) // 2) banner.paste(logo, logo_position)
    # Speichern Sie das resultierende Bild
    banner.save(output_path)

create_banner_with_gradient("..data/logo.png", "..data/banner_with_gradient.png", (255, 0, 0), (0, 0, 255), 200)