from selenium import webdriver
import time
import re
from csv import writer
import pandas as pd
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
store_url = "https://www.aliexpress.com/store/911548204?spm=a2g0o.store_pc_allProduct.pcShopHead_6000579173654.0"
file_name = "Store 2"

# Store links
# https://www.aliexpress.com/store/911548204?spm=a2g0o.store_pc_allProduct.pcShopHead_6000579173654.0
# https://www.aliexpress.com/store/911747985?spm=a2g0o.store_pc_allProduct.pcShopHead_6000712450303.0
# https://www.aliexpress.com/store/5247172?spm=a2g0o.store_pc_allProduct.pcShopHead_373677452.0
# https://www.aliexpress.com/store/910342196?spm=a2g0o.store_pc_allProduct.pcShopHead_1437629253.0
# https://www.aliexpress.com/store/912655354?spm=a2g0o.store_pc_allProduct.pcShopHead_6001934725637.0

driver = webdriver.Chrome("C:\webdrivers\chromedriver.exe")
url = f"{store_url}"
driver.get(url)
driver.maximize_window()

products = driver.find_elements_by_xpath("//a[@class='pc-store-nav-Products']")[0]
products.click()
time.sleep(1)
variance_count = []
prod = []
final_prod = []
def get_product_info():
    prod_url = driver.current_url
    product_qty = driver.find_element_by_xpath("//div[@class='product-quantity-tip']/span")
    # Shipping Cost
    product_shipping = driver.find_element_by_xpath('//div[@class="dynamic-shipping-line dynamic-shipping-titleLayout"]/span/span/strong')    
    product_name = driver.find_element_by_xpath('//div[@class="product-title"]/h1')
    
    try:
        product_cost = driver.find_element_by_xpath('//span[@class="product-price-value"]')
    except:
        product_cost = driver.find_element_by_xpath('//span[@class="uniform-banner-box-price"]')
        
    
    product_data = {'Title': product_name.text,
                    'Product Url': prod_url,
                    'Overall Cost': product_cost.text,
                    'Quantity': product_qty.text,
                    'Shipping': product_shipping.text.replace("-", "to")   
                    }

    prod = [product_name.text, prod_url, product_cost.text, product_qty.text, product_shipping.text.replace("-", "to")]
    fi = len(final_prod)

    time.sleep(1)   

    try:
        variance_set = driver.find_element_by_xpath('//ul[@class="sku-property-list"]').find_elements_by_tag_name("li")
    
        variance_count.append(len(variance_set))
        for k in range(len(variance_set)): 
            variance_set[k].click()

            try:
                variance_name = 'Variance ' + variance_set[k].find_element_by_tag_name("img").get_attribute('alt')  
            except:
                variance_name = 'Variance ' +  variance_set[k].find_element_by_xpath("//div[@class='sku-property-text']/span").text

            variance_cost = product_cost.text
            try:
                variance_image = variance_set[k].find_element_by_tag_name("img").get_attribute('src').replace("jpg_50x50.jpg_.webp", "jpg")
                prod.append(variance_image)
            except:
                pass
            variance_qty = product_qty.text
            prod.append(variance_name)
            prod.append(variance_cost)
            prod.append(variance_qty)
            final_prod.append(prod)
            
            prod = [product_name.text, prod_url, "", "", ""]

    except:
        print("No variance Set")
        prod = [product_name.text, prod_url, product_cost.text, product_qty.text, product_shipping.text.replace("-", "to")]
        final_prod.append(prod)
              

    driver.execute_script("window.scrollTo(0, 1080)")
    time.sleep(2)
    
    
    try:
        product_desc = WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.XPATH, "/html/body/div[7]/div/div[3]/div[3]/div[2]/div[2]/div/div[2]/div[1]/div/div")))
        pattern_size = re.compile("Size")
        product_desc_size = product_desc.find_elements_by_tag_name("div")
        product_desc_imgs = product_desc.find_elements_by_tag_name("img")
        product_desc_span_text = product_desc.find_elements_by_tag_name("span")
        product_desc_text = product_desc.find_elements_by_tag_name("div")
        product_desc_text_p = product_desc.find_elements_by_tag_name("p")
        
        
        for size in product_desc_size:
            match = pattern_size.match(size.text)
            if match:
                product_data = 'Product Size - ' + size.text
                final_prod[fi].append(product_data) 
    
        if(len(product_desc_imgs) > 0):
            for imgs in range(1,len(product_desc_imgs)):
                product_data = 'Product Description Images - ' + product_desc_imgs[imgs].get_attribute("src")
                final_prod[fi].append(f"{product_data}")
                
        if(len(product_desc_text) > 0):
            for text in range(1,30):
                product_data = 'Product Description Text - ' + product_desc_text[text].text
                final_prod[fi].append(f"{product_data}")     
                
        if(len(product_desc_span_text) > 0):
            for span in range(1,30):
                product_data = 'Product Description Text - ' + product_desc_span_text[span].text
                final_prod[fi].append(f"{product_data}")
        
        if(len(product_desc_text_p) > 0):
            for p in range(1,30):
                product_data = 'Product Description Text - ' + product_desc_text_p[p].text
                final_prod[fi].append(f"{product_data}") 
                            
        final_prod.append("")    
        
    except Exception as e:
        pass
        
    final_prod.append("")    
    
    
resultSet = driver.find_element_by_xpath('//ul[@class="items-list util-clearfix"]')
options = resultSet.find_elements_by_tag_name("li")

while True:
     
    check_disabled = driver.find_elements_by_xpath('//span[@class="ui-pagination-next ui-pagination-disabled"]')
    check_next_btn = len(check_disabled)   
    print(check_next_btn)   

    if check_next_btn > 0:
        print("Last Page")
        
        try:
            for option in range(1,len(options)+1):
                try:
                    WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.XPATH, f"/html/body/div[1]/div[5]/div[2]/div/div[2]/div[1]/div/div/div[5]/div/div/ul/li[{option}]/div[2]/h3/a"))).click()
                except:
                    WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.XPATH, f"/html/body/div[2]/div[5]/div[2]/div/div[2]/div[1]/div/div/div[5]/div/div/ul/li[{option}]/div[2]/h3/a"))).click()
                time.sleep(2)
                get_product_info()
                driver.back()
                                
        except:    
            break
        break

    
    if check_next_btn < 1:
        for option in range(1,len(options)+1):
            try:
                WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.XPATH, f"/html/body/div[1]/div[5]/div[2]/div/div[2]/div[1]/div/div/div[5]/div/div/ul/li[{option}]/div[2]/h3/a"))).click()
            except:
                WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.XPATH, f"/html/body/div[2]/div[5]/div[2]/div/div[2]/div[1]/div/div/div[5]/div/div/ul/li[{option}]/div[2]/h3/a"))).click()
            time.sleep(2)
            get_product_info()
            driver.back()
            
        try:
            driver.find_elements_by_xpath('/html/body/div[1]/div[5]/div[2]/div/div[2]/div[1]/div/div/div[6]/div[1]/a')[-1].click()     
            time.sleep(2)
        except:
            pass
            
df = pd.DataFrame(final_prod)
writer = pd.ExcelWriter(f'{file_name}.xlsx', engine='xlsxwriter')
df.to_excel(writer, sheet_name='products', index=False, header=None)
writer.save()

print("Done")