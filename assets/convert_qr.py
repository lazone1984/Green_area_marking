import base64

def convert_image_to_base64(image_path):
    with open(image_path, 'rb') as image_file:
        encoded_string = base64.b64encode(image_file.read())
        return encoded_string.decode('utf-8')

try:
    # 转换两张收款码图片为 base64
    wechat_base64 = convert_image_to_base64('assets/wechat_qr.png')
    alipay_base64 = convert_image_to_base64('assets/alipay_qr.png')

    # 生成要写入的文件内容
    content = f'''# 微信收款码的 base64 数据
WECHAT_QR = """{wechat_base64}"""

# 支付宝收款码的 base64 数据  
ALIPAY_QR = """{alipay_base64}"""
'''

    # 直接写入到 qr_codes.py 文件
    with open('assets/qr_codes.py', 'w', encoding='utf-8') as f:
        f.write(content)

    print("收款码已成功转换并写入 qr_codes.py")
    
except Exception as e:
    print(f"转换过程出错: {str(e)}") 