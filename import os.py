# import os
# import time
# import requests
# import shutil
# from PIL import Image
# from playwright.sync_api import sync_playwright
# from docx import Document
# from docx.shared import Inches

# # 配置参数
# TARGET_URL = "http://www.netbian.com/"
# TEMP_IMG_DIR = os.path.abspath("images")
# CONVERTED_IMG_DIR = os.path.abspath("converted_images")
# OUTPUT_WORD = os.path.abspath("网页内容汇总_02.docx")

# # 初始化文件夹
# for dir_path in [TEMP_IMG_DIR, CONVERTED_IMG_DIR]:
#     if os.path.exists(dir_path):
#         shutil.rmtree(dir_path)
#     os.makedirs(dir_path, exist_ok=True)

# def crawl_content():
#     texts = []
#     img_urls = []
#     cookies = []
    
#     with sync_playwright() as p:
#         browser = p.firefox.launch(headless=False)
#         page = browser.new_page()
        
#         try:
#             page.goto(TARGET_URL)
#             page.wait_for_load_state("networkidle", timeout=60000)
#             time.sleep(10)
            
#             # 滑动加载
#             for i in range(5):
#                 page.evaluate("window.scrollTo(0, document.body.scrollHeight)")
#                 time.sleep(3)
#                 print(f"滑动加载第{i+1}次")
            
#             # 提取文字
#             paragraphs = page.locator("p, .content, .desc").all()
#             for p in paragraphs:
#                 text = p.text_content().strip()
#                 if text and len(text) > 5:
#                     texts.append(text)
#             print(f"提取到{len(texts)}段文字")
            
#             # 提取图片链接（过滤可能的无效格式）
#             imgs = page.locator("img:not([src*='logo']):not([src*='small'])").all()
#             for img in imgs:
#                 src = img.get_attribute("data-src") or img.get_attribute("src")
#                 if src and src.startswith(("http", "https")):
#                     # 优先选择非AVIF格式（如果链接中有格式标识）
#                     if "avif" not in src.lower():
#                         src = src.replace("thumbnail", "original").replace("small", "big")
#                         img_urls.append(src)
#             img_urls = list(set(img_urls))
#             print(f"提取到{len(img_urls)}张图片链接（已过滤AVIF）")
            
#             # 获取Cookie
#             cookies = page.context.cookies()
            
#         finally:
#             browser.close()
    
#     return texts, img_urls, cookies

# def download_images(img_urls, cookies):
#     img_paths = []
#     if not img_urls:
#         print("无图片链接可下载")
#         return img_paths
    
#     cookie_str = "; ".join([f"{c['name']}={c['value']}" for c in cookies])
#     headers = {
#         "User-Agent": "Mozilla/5.0 (X11; Ubuntu; Linux x86_64; rv:136.0) Gecko/20100101 Firefox/136.0",
#         "Cookie": cookie_str,
#         "Referer": TARGET_URL
#     }
    
#     for i, url in enumerate(img_urls, 1):
#         try:
#             print(f"正在下载图片{i}/{len(img_urls)}：{url}")
#             response = requests.get(url, headers=headers, timeout=20, stream=True)
#             response.raise_for_status()
            
#             content_length = int(response.headers.get("Content-Length", 0))
#             if content_length > 0 and content_length < 50 * 1024:  # 过滤小于50KB的图片
#                 print(f"图片过小（{content_length}字节），跳过")
#                 continue
            
#             # 提取并清理文件扩展名
#             url_parts = url.split("?")[0].split(".")
#             ext = url_parts[-1].lower() if len(url_parts) > 1 else "jpg"
#             # 只保留常见图片格式
#             if ext not in ["jpg", "jpeg", "png", "webp", "gif"]:
#                 ext = "jpg"
            
#             img_path = os.path.join(TEMP_IMG_DIR, f"img_{i}.{ext}")
            
#             with open(img_path, "wb") as f:
#                 for chunk in response.iter_content(chunk_size=1024):
#                     if chunk:
#                         f.write(chunk)
            
#             # 验证文件有效性（大小>1KB）
#             if os.path.exists(img_path) and os.path.getsize(img_path) > 1024:
#                 img_paths.append(img_path)
#                 print(f"图片{i}保存成功")
        
#         except Exception as e:
#             print(f"图片{i}下载失败：{str(e)}")
#             continue
    
#     print(f"成功下载{len(img_paths)}/{len(img_urls)}张图片")
#     return img_paths

# def convert_images(img_paths):
#     converted_paths = []
#     for i, img_path in enumerate(img_paths, 1):
#         try:
#             # 尝试打开图片（支持AVIF）
#             with Image.open(img_path) as img:
#                 # 处理透明通道（转为白色背景）
#                 if img.mode in ("RGBA", "LA"):
#                     background = Image.new(img.mode[:-1], img.size, (255, 255, 255))
#                     background.paste(img, img.split()[-1])
#                     img = background
#                 # 转换为RGB模式
#                 if img.mode != "RGB":
#                     img = img.convert("RGB")
#                 # 保存为JPG
#                 converted_path = os.path.join(CONVERTED_IMG_DIR, f"converted_{i}.jpg")
#                 img.save(converted_path, "JPEG", quality=95)
#                 converted_paths.append(converted_path)
#                 print(f"图片{i}转换为JPG成功：{converted_path}")
#         except Exception as e:
#             print(f"图片{i}转换失败：{e}")
#             # 尝试备用方案：直接复制文件（如果是JPG/PNG）
#             ext = os.path.splitext(img_path)[1].lower()
#             if ext in [".jpg", ".jpeg", ".png"]:
#                 converted_path = os.path.join(CONVERTED_IMG_DIR, f"converted_{i}{ext}")
#                 shutil.copy(img_path, converted_path)
#                 converted_paths.append(converted_path)
#                 print(f"图片{i}直接复制备用：{converted_path}")
#     return converted_paths

# def generate_word(texts, converted_img_paths):
#     try:
#         doc = Document()
#         doc.add_heading("网页内容汇总", level=1)
        
#         # 添加文字
#         if texts:
#             doc.add_heading("文字内容", level=2)
#             for i, text in enumerate(texts, 1):
#                 doc.add_paragraph(text)
#                 print(f"添加文字段落{i}")
#             doc.add_page_break()
#         else:
#             doc.add_paragraph("未提取到文字内容")
        
#         # 添加转换后的图片
#         if converted_img_paths:
#             doc.add_heading("图片内容", level=2)
#             for i, img_path in enumerate(converted_img_paths, 1):
#                 if os.path.exists(img_path) and os.path.getsize(img_path) > 1024:
#                     try:
#                         doc.add_picture(img_path, width=Inches(6))
#                         doc.add_paragraph(f"图片{i}")
#                         print(f"添加图片{i}到文档")
#                     except Exception as e:
#                         print(f"图片{i}插入失败：{e}")
#                 else:
#                     print(f"图片{i}文件无效，跳过")
#         else:
#             doc.add_paragraph("未找到有效图片")
        
#         doc.save(OUTPUT_WORD)
#         if os.path.exists(OUTPUT_WORD):
#             print(f"文档保存成功：{OUTPUT_WORD}（大小：{os.path.getsize(OUTPUT_WORD)}字节）")
#         else:
#             print("文档保存失败")
    
#     except Exception as e:
#         print(f"生成Word失败：{e}")

# def main():
#     texts, img_urls, cookies = crawl_content()
#     raw_img_paths = download_images(img_urls, cookies)
#     converted_img_paths = convert_images(raw_img_paths)
#     generate_word(texts, converted_img_paths)
#     print(f"转换后图片目录：{CONVERTED_IMG_DIR}（共{len(os.listdir(CONVERTED_IMG_DIR))}个文件）")

# if __name__ == "__main__":
#     main()



import os
import shutil
import requests
from PIL import Image
from docx import Document
from docx.shared import Inches
from openpyxl import Workbook
from openpyxl.drawing.image import Image as ExcelImage
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.firefox.service import Service
import time

# 配置参数（请确保URL正确）
URL = "https://www.netbian.com"  # 示例：替换为你需要爬取的网页
SAVE_DIR = os.path.join(os.path.expanduser("~"), "桌面", "PC_Test")
ORIGINAL_IMG_DIR = os.path.join(SAVE_DIR, "original_images")
CONVERTED_IMG_DIR = os.path.join(SAVE_DIR, "converted_images")
MIN_IMG_SIZE = 30000
LOAD_TIMES = 5
GECKODRIVER_PATH = ".../geckodriver-v0.36.0-linux64/geckodriver"

# 初始化目录
def init_dirs():
    for dir_path in [SAVE_DIR, ORIGINAL_IMG_DIR, CONVERTED_IMG_DIR]:
        os.makedirs(dir_path, exist_ok=True)
    print(f"工作目录初始化完成: {SAVE_DIR}")

# 滑动加载页面
def load_page(driver):
    try:
        driver.get(URL)
        print(f"已成功访问网页：{URL}")
        time.sleep(3)  # 延长初始加载时间
    except Exception as e:
        print(f"访问网页失败：{str(e)}")
        return driver
    
    for i in range(LOAD_TIMES):
        try:
            driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")
            print(f"滑动加载第{i+1}次")
            time.sleep(3)  # 延长滑动后等待时间
        except Exception as e:
            print(f"滑动加载失败：{str(e)}")
            break
    return driver

# 提取文字段落（适配网页结构）
def extract_text(driver):
    try:
        # 尝试多种标签提取文字（根据网页调整）
        selectors = [By.TAG_NAME, "p", By.TAG_NAME, "div", By.TAG_NAME, "span"]
        text_list = []
        for i in range(0, len(selectors), 2):
            by = selectors[i]
            tag = selectors[i+1]
            elements = driver.find_elements(by, tag)
            text_list.extend([e.text.strip() for e in elements if e.text.strip()])
        # 去重
        text_list = list(set(text_list))
        print(f"提取到{len(text_list)}段文字")
        return text_list
    except Exception as e:
        print(f"提取文字失败：{str(e)}")
        return []

# 提取图片链接
def extract_images(driver):
    try:
        img_tags = driver.find_elements(By.TAG_NAME, "img")
        img_urls = []
        for img in img_tags:
            url = img.get_attribute("src")
            if url:
                # 补全相对路径URL
                if url.startswith("//"):
                    url = "https:" + url
                elif url.startswith("/"):
                    url = f"{URL}{url}"
                if "avif" not in url.lower():
                    img_urls.append(url)
        # 去重
        img_urls = list(set(img_urls))
        print(f"提取到{len(img_urls)}张图片链接（已过滤AVIF）")
        return img_urls
    except Exception as e:
        print(f"提取图片链接失败：{str(e)}")
        return []

# 下载图片
def download_images(img_urls):
    saved_paths = []
    for i, url in enumerate(img_urls, 1):
        try:
            print(f"正在下载图片{i}/{len(img_urls)}：{url}")
            headers = {
                "User-Agent": "Mozilla/5.0 (X11; Ubuntu; Linux x86_64; rv:129.0) Gecko/20100101 Firefox/129.0",
                "Referer": URL  # 添加来源页，避免反爬
            }
            response = requests.get(url, headers=headers, timeout=15)
            response.raise_for_status()
            
            if len(response.content) < MIN_IMG_SIZE:
                print(f"图片过小（{len(response.content)}字节），跳过")
                continue
            
            ext = os.path.splitext(url.split('?')[0])[-1] or ".jpg"
            img_name = f"img_{i}{ext}"
            original_path = os.path.join(ORIGINAL_IMG_DIR, img_name)
            with open(original_path, "wb") as f:
                f.write(response.content)
            
            saved_paths.append(original_path)
            print(f"图片{i}保存成功")
        except Exception as e:
            print(f"图片{i}下载失败：{str(e)}")
    print(f"成功下载{len(saved_paths)}/{len(img_urls)}张图片")
    return saved_paths

# 转换图片
def convert_images(original_paths):
    converted_paths = []
    for i, original_path in enumerate(original_paths, 1):
        try:
            with Image.open(original_path) as img:
                if img.mode in ("RGBA", "P"):
                    img = img.convert("RGB")
                converted_path = os.path.join(CONVERTED_IMG_DIR, f"converted_{i}.jpg")
                img.save(converted_path, "JPEG")
                converted_paths.append(converted_path)
                print(f"图片{i}转换为JPG成功：{converted_path}")
        except Exception as e:
            converted_path = os.path.join(CONVERTED_IMG_DIR, f"converted_{i}{os.path.splitext(original_path)[-1]}")
            shutil.copy2(original_path, converted_path)
            converted_paths.append(converted_path)
            print(f"图片{i}转换失败：{str(e)}，已复制原始文件备用")
    return converted_paths

# 保存为Word
def save_to_word(texts, images):
    try:
        doc = Document()
        for i, text in enumerate(texts, 1):
            doc.add_paragraph(text)
            print(f"添加文字段落{i}")
        for i, img_path in enumerate(images, 1):
            try:
                doc.add_picture(img_path, width=Inches(5))
                print(f"添加图片{i}到文档")
            except Exception as e:
                print(f"图片{i}插入失败：{str(e)}")
        doc_path = os.path.join(SAVE_DIR, "网页内容汇总.docx")
        doc.save(doc_path)
        print(f"文档保存成功：{doc_path}")
    except Exception as e:
        print(f"保存Word失败：{str(e)}")

# 保存为Excel
def save_to_excel(texts, images):
    try:
        wb = Workbook()
        ws = wb.active
        ws.title = "网页内容汇总"
        ws.append(["文字内容"])
        for i, text in enumerate(texts, 1):
            ws.cell(row=i+1, column=1, value=text)
            ws.row_dimensions[i+1].height = 40
            print(f"Excel添加文字段落{i}")
        current_row = len(texts) + 3
        ws.cell(row=current_row-1, column=1, value="图片内容")
        for i, img_path in enumerate(images, 1):
            try:
                img = ExcelImage(img_path)
                img.width = 400
                img.height = 300
                ws.add_image(img, f"A{current_row}")
                ws.row_dimensions[current_row].height = 250
                current_row += 10
                print(f"Excel添加图片{i}")
            except Exception as e:
                print(f"Excel图片{i}插入失败：{str(e)}")
        excel_path = os.path.join(SAVE_DIR, "网页内容汇总.xlsx")
        wb.save(excel_path)
        print(f"Excel保存成功：{excel_path}")
    except Exception as e:
        print(f"保存Excel失败：{str(e)}")

# 主函数
def main():
    init_dirs()
    
    if not os.path.exists(GECKODRIVER_PATH):
        print(f"错误：驱动文件不存在 - {GECKODRIVER_PATH}")
        return
    
    try:
        # 简化驱动配置，使用已验证的正常驱动
        service = Service(executable_path=GECKODRIVER_PATH)
        driver = webdriver.Firefox(service=service)
        # 最大化窗口，确保内容加载完整
        driver.maximize_window()
        # 设置超时时间
        driver.set_page_load_timeout(30)
        driver.implicitly_wait(10)  # 元素查找超时
    except Exception as e:
        print(f"启动火狐浏览器失败：{str(e)}")
        return
    
    try:
        driver = load_page(driver)
        texts = extract_text(driver)
        img_urls = extract_images(driver)
        original_img_paths = download_images(img_urls)
        converted_img_paths = convert_images(original_img_paths)
        save_to_word(texts, converted_img_paths)
        save_to_excel(texts, converted_img_paths)
        print(f"原始图片目录：{ORIGINAL_IMG_DIR}（共{len(os.listdir(ORIGINAL_IMG_DIR))}个文件）")
        print(f"转换后图片目录：{CONVERTED_IMG_DIR}（共{len(os.listdir(CONVERTED_IMG_DIR))}个文件）")
    finally:
        # 确保浏览器关闭
        try:
            driver.quit()
        except:
            pass

if __name__ == "__main__":
    main()


