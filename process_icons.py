from PIL import Image
import os

def add_white_background(input_path, output_path):
    # 画像を開く
    img = Image.open(input_path)
    
    # RGBAモードでない場合は変換
    if img.mode != 'RGBA':
        img = img.convert('RGBA')
    
    # 同じサイズの白背景を作成
    background = Image.new('RGBA', img.size, (255, 255, 255, 255))
    
    # 画像を白背景に合成
    white_image = Image.alpha_composite(background, img)
    
    # 保存
    white_image.save(output_path, 'PNG')

# MyIcon.iconset内の全てのファイルを処理
iconset_path = 'MyIcon.iconset'
for filename in os.listdir(iconset_path):
    if filename.endswith('.png'):
        input_path = os.path.join(iconset_path, filename)
        output_path = os.path.join(iconset_path, filename)
        add_white_background(input_path, output_path)

print("全てのアイコンの背景を白に変更しました。")