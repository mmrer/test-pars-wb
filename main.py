from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
from openpyxl import Workbook
import time, random, json, re


# настройка браузера(без детекции)
def setup_driver():
    options = webdriver.ChromeOptions()
    options.add_argument('--disable-blink-features=AutomationControlled')
    options.add_experimental_option("excludeSwitches", ["enable-automation"])
    options.add_experimental_option('useAutomationExtension', False)
    options.add_argument('user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36')
    
    service = Service(ChromeDriverManager().install())
    driver = webdriver.Chrome(service=service, options=options)
    return driver

# парсинг страницы поиска и фильтрация по рейтингу
def parse_search_page(driver, url, min_rating=4.5):
    print("открытие страницы...")
    driver.get(url)
    time.sleep(random.uniform(2, 4))
    
    # прокрутка страницы
    for i in range(3):
        driver.execute_script(f"window.scrollTo(0, {i * 500});")
        time.sleep(random.uniform(0.5, 1.5))
    
    driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")
    time.sleep(random.uniform(1, 2))
    
    # поиск всех карточек товаров
    cards = driver.find_elements(By.CSS_SELECTOR, "article.product-card.j-card-item")
    print(f"найдено карточек: {len(cards)}")
    
    # фильтрация по рейтингу
    filtered_articles = []
    for card in cards:
        try:
            rating_element = card.find_element(By.CSS_SELECTOR, "span.address-rate-mini")
            rating_text = rating_element.text.replace(',', '.')
            rating = float(rating_text)
            
            if rating >= min_rating:
                article = card.get_attribute("data-nm-id")
                if article:
                    filtered_articles.append(article)
                    print(f"+ Артикул {article} - рейтинг {rating}")
        except:
            continue
    
    print(f"\nвсего товаров с рейтингом >= {min_rating}: {len(filtered_articles)}")
    return filtered_articles

# создание словаря для хранения данных о товаре
def create_product_dict(article, product_url):
    return {
        'Ссылка на товар': product_url,
        'Артикул': article,
        'Название': '',
        'Цена': '',
        'Ссылки на изображения': '',
        'Описание': '',
        'Характеристики': '',
        'Название продавца': '',
        'Ссылка на продавца': '',
        'Размеры товара': '',
        'Остатки по товару': '',
        'Рейтинг': '',
        'Количество отзывов': ''
    }

# парсинг названия товара
def parse_product_name(driver, product_info, article):
    try:
        title_element = WebDriverWait(driver, 5).until(
            EC.presence_of_element_located((By.CSS_SELECTOR, "h3.productTitle--J2W7I, h3.mo-typography_variant_title3"))
        )
        product_info['Название'] = title_element.text.strip()
    except:
        print(f"  не удалось найти название для {article}")

# парсинг цены товара
def parse_product_price(driver, product_info, article):
    try:
        price_element = WebDriverWait(driver, 5).until(
            EC.presence_of_element_located((By.CSS_SELECTOR, "h2.mo-typography_color_danger, ins.priceBlockFinalPrice--iToZR"))
        )
        price_text = price_element.text.strip()
        price_text = price_text.replace('\xa0', '').replace('₽', '').strip()
        product_info['Цена'] = price_text
    except:
        print(f"  не удалось найти цену для {article}")

# парсинг рейтинга и количества отзывов
def parse_product_rating(driver, product_info, article):
    try:
        rating_element = WebDriverWait(driver, 5).until(
            EC.presence_of_element_located((By.CSS_SELECTOR, "span.productReviewRating--gQDQG"))
        )
        rating_text = rating_element.text.strip()
        
        parts = rating_text.split('·')
        if len(parts) >= 2:
            product_info['Рейтинг'] = parts[0].strip()
            reviews_text = parts[1].strip()
            reviews_match = re.search(r'(\d+)', reviews_text)
            product_info['Количество отзывов'] = reviews_match.group(1) if reviews_match else '0'
        else:
            numbers = re.findall(r'(\d+)', rating_text)
            if len(numbers) >= 1:
                product_info['Рейтинг'] = numbers[0]
            if len(numbers) >= 2:
                product_info['Количество отзывов'] = numbers[1]
    except:
        print(f"  не удалось найти рейтинг и отзывы для {article}")

# парсинг изображений товара
def parse_product_images(driver, product_info, article):
    images = []
    try:
        WebDriverWait(driver, 5).until(
            EC.presence_of_element_located((By.CSS_SELECTOR, "div.miniaturesWrapper--PF0rM, div.swiper-wrapper.miniaturesWrapper--PF0rM"))
        )
        time.sleep(0.5)
        img_elements = driver.find_elements(By.CSS_SELECTOR, "div.miniaturesWrapper--PF0rM img, div.swiper-wrapper.miniaturesWrapper--PF0rM img")
        for img in img_elements:
            img_url = img.get_attribute("src")
            if img_url and img_url.startswith("http"):
                img_url = img_url.replace("/c246x328/", "/c516x688/")
                if img_url not in images:
                    images.append(img_url)
        product_info['Ссылки на изображения'] = ', '.join(images)
    except:
        print(f"  не удалось найти изображения для {article}")
    return images

# парсинг информации о продавце
def parse_seller_info(driver, product_info, article):
    try:
        seller_link = WebDriverWait(driver, 5).until(
            EC.presence_of_element_located((By.CSS_SELECTOR, "a.sellerInfoButtonLink--RoLBz"))
        )
        seller_href = seller_link.get_attribute("href")
        if seller_href:
            product_info['Ссылка на продавца'] = seller_href
        
        try:
            seller_name_element = seller_link.find_element(By.CSS_SELECTOR, "span.sellerInfoNameDefaultText--qLwgq")
            product_info['Название продавца'] = seller_name_element.text.strip()
        except:
            print(f"  не удалось найти название продавца для {article}")
    except:
        print(f"  не удалось найти информацию о продавце для {article}")

# парсинг размеров и остатков товара
def parse_sizes_and_stock(driver, product_info, article):
    try:
        sizes_list = WebDriverWait(driver, 5).until(
            EC.presence_of_element_located((By.CSS_SELECTOR, "ul.sizesList--EwFfe"))
        )
        time.sleep(0.5)
        size_items = sizes_list.find_elements(By.CSS_SELECTOR, "li.sizesListItem--QcbQx")
        
        sizes = []
        total_stock = 0
        
        for size_item in size_items:
            if "sizeDisabled" in size_item.get_attribute("class"):
                continue
            
            try:
                # Ищем размер внутри конкретного элемента size_item
                size_element = size_item.find_element(By.CSS_SELECTOR, "span.sizesListSize--NUoNC")
                size_text = size_element.text.strip()
                
                time.sleep(0.3)
                actions = ActionChains(driver)
                actions.move_to_element(size_item).perform()
                time.sleep(0.8)
                
                stock_info = get_stock_from_tooltip(driver)
                if stock_info != "-":
                    total_stock += int(stock_info)
                
                sizes.append(f"{size_text}:{stock_info}")
            except:
                continue
        
        product_info['Размеры товара'] = ', '.join(sizes)
        product_info['Остатки по товару'] = str(total_stock) if total_stock > 0 else '-'
    except Exception as e:
        print(f"  не удалось найти размеры для {article}: {e}")

# получение остатков из подсказки
def get_stock_from_tooltip(driver):
    stock_info = "-"
    tooltip_text = None
    
    try:
        tooltip_selectors = [
            "[class*='tooltip']", "[class*='popover']", "[class*='hint']",
            "[class*='stock']", "[role='tooltip']", "[class*='Tooltip']", "[class*='Popover']"
        ]
        
        for selector in tooltip_selectors:
            try:
                tooltips = driver.find_elements(By.CSS_SELECTOR, selector)
                for tooltip in tooltips:
                    if tooltip.is_displayed():
                        tooltip_text = tooltip.text
                        break
                if tooltip_text:
                    break
            except:
                continue
        
        if not tooltip_text:
            try:
                js_code = """
                var tooltips = document.querySelectorAll('[class*="tooltip"], [class*="popover"], [class*="hint"], [class*="stock"]');
                for (var i = 0; i < tooltips.length; i++) {
                    var el = tooltips[i];
                    if (el.offsetParent !== null) {
                        return el.textContent;
                    }
                }
                return null;
                """
                tooltip_text = driver.execute_script(js_code)
            except:
                pass
        
        if tooltip_text and "осталось менее" in tooltip_text.lower():
            stock_match = re.search(r'осталось\s+менее\s+(\d+)', tooltip_text, re.IGNORECASE)
            if stock_match:
                stock_info = stock_match.group(1)
    except:
        pass
    
    return stock_info

# парсинг описания и характеристик товара
def parse_description_and_characteristics(driver, product_info, article):
    try:
        detail_button = WebDriverWait(driver, 5).until(
            EC.element_to_be_clickable((By.CSS_SELECTOR, "button.btnDetail--im7UR"))
        )
        time.sleep(0.5)
        detail_button.click()
        
        try:
            WebDriverWait(driver, 5).until(
                EC.presence_of_element_located((By.CSS_SELECTOR, "section#section-description, section[data-testid='product_additional_information']"))
            )
        except:
            time.sleep(2)
        
        # Описание
        try:
            desc_section = WebDriverWait(driver, 3).until(
                EC.presence_of_element_located((By.CSS_SELECTOR, "section#section-description"))
            )
            desc_text = desc_section.find_element(By.CSS_SELECTOR, "p.descriptionText--Jq9n2")
            product_info['Описание'] = desc_text.text.strip()
        except:
            print(f"  не удалось найти описание для {article}")
        
        # Характеристики
        try:
            characteristics_section = WebDriverWait(driver, 3).until(
                EC.presence_of_element_located((By.CSS_SELECTOR, "section[data-testid='product_additional_information']"))
            )
            time.sleep(0.5)
            tables = characteristics_section.find_elements(By.CSS_SELECTOR, "table.table--tSF0X")
            
            characteristics = {}
            for table in tables:
                try:
                    caption = table.find_element(By.CSS_SELECTOR, "caption.caption--gsljv")
                    section_name = caption.text.strip()
                except:
                    section_name = "Без названия"
                
                rows = table.find_elements(By.CSS_SELECTOR, "tbody tr")
                section_data = {}
                
                for row in rows:
                    try:
                        key_element = row.find_element(By.CSS_SELECTOR, "th.cellKey--eGe6N")
                        value_element = row.find_element(By.CSS_SELECTOR, "td.cellValue--hHBJB")
                        section_data[key_element.text.strip()] = value_element.text.strip()
                    except:
                        continue
                
                if section_data:
                    characteristics[section_name] = section_data
            
            product_info['Характеристики'] = json.dumps(characteristics, ensure_ascii=False, indent=2)
        except:
            print(f"  не удалось найти характеристики для {article}")
    except:
        print(f"  не удалось открыть панель характеристик для {article}")

# вывод информации о товаре
def print_product_info(product_info, images):
    print(f"\nАртикул: {product_info['Артикул']}")
    print(f"+ Название: {product_info['Название']}")
    print(f"+ Цена: {product_info['Цена']}")
    print(f"+ Рейтинг: {product_info['Рейтинг']}")
    print(f"+ Количество отзывов: {product_info['Количество отзывов']}")
    print(f"+ Ссылка на товар: {product_info['Ссылка на товар']}")
    print(f"+ Изображений: {len(images)}")
    print(f"+ Продавец: {product_info['Название продавца']}")
    print(f"+ Ссылка на продавца: {product_info['Ссылка на продавца']}")
    print(f"+ Размеры: {product_info['Размеры товара']}")
    print(f"+ Остатки: {product_info['Остатки по товару']}")
    print(f"+ Описание: {'Да' if product_info['Описание'] else 'Нет'}")
    print(f"+ Характеристики: {'Да' if product_info['Характеристики'] else 'Нет'}")

# парсинг страницы товара
def parse_product_page(driver, article):
    product_url = f"https://www.wildberries.ru/catalog/{article}/detail.aspx"
    
    driver.get(product_url)
    try:
        WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.CSS_SELECTOR, "h3.productTitle--J2W7I, h3.mo-typography_variant_title3"))
        )
    except:
        time.sleep(3)
    
    product_info = create_product_dict(article, product_url)
    
    parse_product_name(driver, product_info, article)
    parse_product_price(driver, product_info, article)
    parse_product_rating(driver, product_info, article)
    parse_product_images(driver, product_info, article)
    parse_seller_info(driver, product_info, article)
    parse_sizes_and_stock(driver, product_info, article)
    parse_description_and_characteristics(driver, product_info, article)
    
    return product_info

# сохранение каталога в excel
def save_to_excel(products_data, filename='wildberries_catalog.xlsx'):
    if not products_data:
        print("нет данных для сохранения")
        return
    
    wb = Workbook()
    ws = wb.active
    ws.title = "Товары wb"
    
    headers = ['Ссылка на товар', 'Артикул', 'Название', 'Цена', 'Описание', 
              'Ссылки на изображения', 'Характеристики', 'Название продавца', 
              'Ссылка на продавца', 'Размеры товара', 'Остатки по товару', 
              'Рейтинг', 'Количество отзывов']
    
    ws.append(headers)
    
    for product in products_data:
        row = [product.get(header, '') for header in headers]
        ws.append(row)
    
    wb.save(filename)
    print(f"\n✓ Данные сохранены в {filename}")
    print(f"  Всего товаров: {len(products_data)}")


def main():
    driver = setup_driver()
    
    url = "https://www.wildberries.ru/catalog/0/search.aspx?page=1&sort=popular&search=%D0%BF%D0%B0%D0%BB%D1%8C%D1%82%D0%BE+%D0%B8%D0%B7+%D0%BD%D0%B0%D1%82%D1%83%D1%80%D0%B0%D0%BB%D1%8C%D0%BD%D0%BE%D0%B9+%D1%88%D0%B5%D1%80%D1%81%D1%82%D0%B8&priceU=980000%3B1000000&f14177451=15000203"
    
    filtered_articles = parse_search_page(driver, url)
    
    products_data = []
    print("\nоткрытие страниц товаров...")
    for i, article in enumerate(filtered_articles, 1):
        print(f"[{i}/{len(filtered_articles)}] ", end="")
        product_info = parse_product_page(driver, article)
        products_data.append(product_info)
        print("+")
    
    time.sleep(2)
    driver.quit()
    
    # сохранение в таблицу
    print("сохранение данных...")
    save_to_excel(products_data, 'wildberries_catalog.xlsx')
    
    return products_data


if __name__ == "__main__":
    products_data = main()