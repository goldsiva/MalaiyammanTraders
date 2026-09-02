import pandas as pd
import os
import re

# --- Cleaning helpers ---
def clean_numeric(val):
    if pd.isna(val):
        return 0
    cleaned = re.sub(r"[^0-9.]", "", str(val))
    try:
        return float(cleaned)
    except ValueError:
        return 0

def clean_text(val):
    if pd.isna(val):
        return ""
    return re.sub(r"[^\w\s\-\(\)/]", "", str(val)).strip()

# --- Load Excel ---
df = pd.read_excel("products.xlsx")
df.columns = [c.strip().split("(")[0].strip() for c in df.columns]
df = df.fillna("")

# --- Apply cleaning ---
if "Selling Price" in df.columns:
    df["Selling Price"] = df["Selling Price"].apply(clean_numeric)
if "MRP" in df.columns:
    df["MRP"] = df["MRP"].apply(clean_numeric)
if "Product Name" in df.columns:
    df["Product Name"] = df["Product Name"].apply(clean_text)
if "Size" in df.columns:
    df["Size"] = df["Size"].apply(clean_text)
if "Product Description" in df.columns:
    df["Product Description"] = df["Product Description"].apply(clean_text)

output_dir = "."
os.makedirs(output_dir, exist_ok=True)

# --- Build categories menu ---
categories = sorted(df["Product Category"].unique())
categories_menu = "<li><a href='#'>Categories</a><ul>\n"
for cat in categories:
    filename = cat.lower().replace(" ", "") + ".html"
    categories_menu += f"<li><a href='{filename}'>{cat}</a></li>\n"
categories_menu += "</ul></li>"

nav_html = f"""
<nav>
  <ul>
    <li><a href="index.html">Home</a></li>
    <li><a href="about.html">About Us</a></li>
    <li><a href="products.html">Products</a></li>
    <li><a href="contact.html">Contact Us</a></li>
    {categories_menu}
  </ul>
</nav>
"""

# --- Keep your original header/footer definitions ---
header_html = """
<header class="site-header">
  <div class="header-left">
    <a href="index.html"><img src="logo.png" alt="Malaiyamman Traders Logo" class="logo"></a>
  </div>
  <div class="header-center">
    <h1>Malaiyamman Traders</h1>
    <p>Crackers & Fireworks</p>
  </div>
  <div class="header-right">
    <a href="tel:+919843611870" class="contact-icon"><i class="fa-solid fa-mobile-screen-button"></i></a>
    <a href="https://facebook.com/malaiyammantraders" target="_blank" class="contact-icon"><i class="fa-brands fa-facebook"></i></a>
    <a href="https://instagram.com/malaiyammantraders" target="_blank" class="contact-icon"><i class="fa-brands fa-instagram"></i></a>
    <div id="cart-summary" class="cart-summary">🛍️ View Cart (0) - ₹0.00</div>
  </div>
</header>
"""

footer_html = """
<footer class="site-footer">
  <div class="footer-nav">
    <a href="index.html">Home</a>
    <a href="about.html">About Us</a>
    <a href="products.html">Products</a>
    <a href="contact.html">Contact Us</a>
  </div>
  <p>&copy; 2026 Malaiyamman Traders. All rights reserved.</p>
  <div class="footer-icons">
    <a href="tel:+919843611870" class="contact-icon"><i class="fa-solid fa-mobile-screen-button"></i></a>
    <a href="https://facebook.com/malaiyammantraders" target="_blank" class="contact-icon"><i class="fa-brands fa-facebook"></i></a>
    <a href="https://instagram.com/malaiyammantraders" target="_blank" class="contact-icon"><i class="fa-brands fa-instagram"></i></a>
  </div>
</footer>
"""

# --- Generate product table ---
def generate_table(group):
    table = '<table class="product-table sortable">\n'
    table += "<tr><th>Product Name</th><th>MRP</th><th>Selling Price</th><th>Description</th><th>Size/Variant</th><th>Action</th></tr>\n"
    for _, row in group.iterrows():
        table += f"""
        <tr>
          <td>{row['Product Name']}</td>
          <td>{row.get('MRP','')}</td>
          <td>{row.get('Selling Price','')}</td>
          <td>{row.get('Product Description','')}</td>
          <td>{row.get('Size','')}</td>
          <td>
            <button class="add-to-cart"
                    data-name="{row['Product Name']}"
                    data-variant="{row.get('Size','')}"
                    data-price="{row.get('Selling Price','0')}">
              <i class="fa fa-cart-plus"></i> Add to Cart
            </button>
          </td>
        </tr>
        """
    table += "</table>\n"
    return table

shared_js = """
<script src="style.js"></script>
<script src="script.js"></script>
"""

shared_css = """

"""

# --- Generate products.html ---
main_html = f"""
<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<title>Products - Malaiyamman Traders</title>
<link rel="stylesheet" href="style.css">
<link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.5.0/css/all.min.css">
{shared_css}
</head>
<body>
{header_html}
{nav_html}
<section>
<h2>Explore Our Categories</h2>
<input type="text" id="searchInput" placeholder="Search products by name...">
"""

for category, group in df.groupby("Product Category"):
    main_html += f'<h3 class="category-title">{category}</h3>\n'
    main_html += generate_table(group)

main_html += f"""
</section>
{footer_html}
{shared_js}
</body>
<div id="toast-container"></div>
</html>
"""

with open(os.path.join(output_dir, "products.html"), "w", encoding="utf-8") as f:
    f.write(main_html)

# --- Generate category pages ---
for category, group in df.groupby("Product Category"):
    filename = category.lower().replace(" ", "") + ".html"
    cat_html = f"""
<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<title>{category} - Malaiyamman Traders</title>
<link rel="stylesheet" href="style.css">
<link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.5.0/css/all.min.css">
{shared_css}
</head>
<body>
{header_html}
{nav_html}
<section>
<h2>{category} Collection</h2>
{generate_table(group)}
</section>
{footer_html}
{shared_js}
<div id="toast-container"></div>
</body>
</html>
"""
    with open(os.path.join(output_dir, filename), "w", encoding="utf-8") as f:
        f.write(cat_html)
