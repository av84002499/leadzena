from bs4 import BeautifulSoup
import openpyxl

# Read the HTML file
with open('Amazon.html', 'r', encoding='utf-8') as file:
    html_content = file.read()

# Create a BeautifulSoup object
soup = BeautifulSoup(html_content, 'html.parser')

# Find all divs with the specified class
divs = soup.find_all('div', class_='s-card-container s-overflow-hidden aok-relative puis-wide-grid-style puis-wide-grid-style-t3 puis-include-content-margin puis puis-v1aj7nq8vmj30z2oahbw0jjbcwc s-latency-cf-section s-card-border')

# Initialize lists to store the extracted data
product_names = []
product_prices = []
product_reviews = []

# Loop through each div and extract the required information
for div in divs:
    # Find the product name
    product_name_element = div.find('span', class_='a-size-medium a-color-base a-text-normal')
    product_name = product_name_element.text.strip() if product_name_element else ''
    product_names.append(product_name)

    # Find the product price
    product_price_element = div.find('span', class_='a-price-whole')
    product_price = product_price_element.text.strip() if product_price_element else ''
    product_prices.append(product_price)

    # Find the product reviews
    product_review_element = div.find('span', class_='a-declarative')
    product_review = product_review_element.text.strip() if product_review_element else ''
    product_reviews.append(product_review)

# Create an Excel file and write the data
workbook = openpyxl.Workbook()
sheet = workbook.active
sheet.append(['Product Name', 'Product Price', 'Product Reviews'])

for name, price, review in zip(product_names, product_prices, product_reviews):
    sheet.append([name, price, review])

workbook.save('Amazon_Products.xlsx')
