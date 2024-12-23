from PIL import Image as PILImage  # 使用PIL來處
from openpyxl import Workbook
from openpyxl.drawing.image import Image

def insert_img(ws:Workbook,start_row:int,image_path, target_height):
    # 讀取圖片
    img = PILImage.open(image_path)
    width, height = img.size
    result_img = None
    if target_height < height:
        # 目標高度小於原圖高度，裁剪圖片
        cropped_img = img.crop((0, 0, width, target_height))
        result_img = cropped_img
    elif target_height > height:
        # 目標高度大於原圖高度，延伸圖片
        extra_height = target_height - height

        # 選取圖片底部一小部分作為紋理進行延伸（選擇高度10%的部分）
        texture_height = max(1, height // 10)  # 保證最小高度為1
        texture = img.crop((0, height - texture_height, width, height))

        # 平鋪紋理，使其達到所需的延伸高度
        num_repeats = (extra_height // texture_height) + 1
        extended_texture = PILImage.new("RGB", (width, extra_height))

        # 將紋理重複填充
        for i in range(num_repeats):
            extended_texture.paste(texture, (0, i * texture_height))

        # 截取到正好需要的高度
        extended_texture = extended_texture.crop((0, 0, width, extra_height))

        # 將原圖與延伸的部分拼接起來
        final_img = PILImage.new("RGB", (width, target_height))
        final_img.paste(img, (0, 0))
        final_img.paste(extended_texture, (0, height))
        result_img = final_img
    else:
        # 如果高度相同，不做處理
        result_img = img
    result_img.save('temp.png')
    insert_img = Image('temp.png')
    ws.add_image(insert_img, f'C{start_row}')

