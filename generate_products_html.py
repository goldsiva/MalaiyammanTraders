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

# --- Safe ID generator ---
def make_id_safe(text):
    return re.sub(r'[^A-Za-z0-9_-]', '_', str(text)).strip('_')

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
# sorted
categories = (df["Product Category"].unique())
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
    <li><a href="products.html"><span class="blink_text">Estimate Now</span></a></li>
    <li><a href="contact.html">Contact Us</a></li>
    {categories_menu}
  </ul>
</nav>
"""

# --- Header/Footer ---
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
    <div id="cart-summary" class="cart-summary" onclick="gotoCart()" style="cursor:pointer;">
      🛍️ Place Order
    </div>
  </div>
</header>
<div class="open-banner">
  <p>🕒 Our shop is open <strong>24/7 – 365 days</strong> for sale!</p>
</div>
"""

footer_html = """
<!-- Footer -->
  <footer class="site-footer footbg footpad footlink">
  <div class="container">
    <div class="row">
      
      <!-- About -->
      <div class="col-lg-4 col-md-4">
        <h1 class="gilroy pb-2 heading2 clr text-white">Malaiyamman Traders</h1>
        <p class="roboto text-white smallfnt">
          Premium Crackers & Fireworks for every occasion. Wide range including Sparklers, Rockets, Flower Pots, and more.
        </p>
      </div>
      
      <!-- Quick Links -->
      <div class="col-lg-4 col-md-4">
        <h1 class="gilroy pb-3 text-white heading2">Quick Links</h1>
        <ul class="fullpad">
          <li class="roboto pb-2"><a href="index.html" class="text-white clr"><i class="fa fa-angle-right"></i> Home</a></li>
          <li class="roboto pb-2"><a href="about.html" class="text-white clr"><i class="fa fa-angle-right"></i> About</a></li>
          <li class="roboto pb-2"><a href="products.html" class="text-white clr"><i class="fa fa-angle-right"></i> Products</a></li>
          <li class="roboto pb-2"><a href="contact.html" class="text-white clr"><i class="fa fa-angle-right"></i> Contact</a></li>
        </ul>
      </div>
      
      <!-- Contact -->
      <div class="col-lg-4 col-md-4">
        <h1 class="gilroy heading2 pb-2 clr text-white">Contact Us</h1>
        <ul class="fullpad">
          <li class="roboto pb-2">
            <i class="fa fa-location-arrow text-white"></i>
            <div class="text1 text-white">
             1/4728/6, Rosalpatti, Thanga Mani Colony, Rosalpatti, Tamil Nadu 626001
            </div>
          </li>
          <li class="roboto pb-2">
            <i class="fa fa-whatsapp text-white"></i>
            <div class="text1 para">
              <a href="https://wa.me/919843611870" class="text-white">+91 9843611870</a>
            </div>
          </li>
          <li class="roboto pb-2">
            <i class="fa fa-envelope text-white"></i>
            <div class="text1 para">
              <a href="mailto:malaiyammantraders@gmail.com" class="text-white">malaiyammantraders@gmail.com</a>
            </div>
          </li>
        </ul>
      </div>
      
    </div> <!-- end row -->
    
    <!-- Bottom row -->
    <div class="row">
      <div class="col-md-6 pt-4">
        <p class="roboto text-white">© 2026 Malaiyamman Traders – All rights reserved</p>
      </div>
      <div class="col-md-6 text-right pt-4" style="text-align: right;">
        <i>Designed by <a href="#" target="_blank" style="color:#9e9e9e; text-decoration:none;">Malaiyamman Traders</a></i>
      </div>
    </div>
    
  </div>
</footer>
"""

# --- Generate single unified table ---
def generate_table(df):
    table = """        
    <table class="product-table sortable">
    <tr>
      <th>Image</th>
      <th>Product Name</th>
      <th>MRP</th>
      <th>Selling Price</th>
      <th>Content</th>
      <th>Size/Variant</th>
      <th>Quantity</th>
      <th>Total</th>
    </tr>
    """
    for category, group in df.groupby("Product Category", sort=False):
        table += f"""
        <tr>
          <td colspan="8" style="text-align:center;font-weight:bold;background:#f0f0f0;color: brown;font-size: 25px;">
            {category}
          </td>
        </tr>
        """
        for _, row in group.iterrows():
            name = row['Product Name']
            variant = row.get('Size','')
            desc = row.get('Product Description','')
            price = row.get('Selling Price','0')
            order = row.get('Order','0')
            img_file = f"{order}.jpg"

            safe_name = make_id_safe(name)
            safe_variant = make_id_safe(variant)

            table += f"""
            <tr>
              <td><img src="images/{img_file}" style="max-width:30px; max-height:40px;"></td>
              <td class="product-name">{name}</td>
              <td>{row.get('MRP','')}</td>
              <td>{price}</td>
              <td>{desc}</td>
              <td>{variant}</td>
              <td>
                <button class="minus remove-from-cart" data-name="{name}" data-variant="{variant}" data-price="{price}"
                    onclick="updateQty('{name}','{variant}',-1)" style="background:red;color:white;"><i class="fa fa-minus" aria-hidden="true"></i></button>
                <input class="num-pallets-input" id="qty-{safe_name}-{safe_variant}" type="text" value="0" data-price="{price}"
                       oninput="validateQty(this,'{name}','{variant}')" >
                <button class="add add-to-cart" data-name="{name}" data-variant="{variant}" data-price="{price}"
                    onclick="updateQty('{name}','{variant}',1)" style="background:green;color:white;"><i class="fa fa-plus" aria-hidden="true"></i></button>
              </td>
              <td id="total-{safe_name}-{safe_variant}">0.00</td>
            </tr>
            """
    table += """
    <tfoot>
      <tr>
        <td colspan="7" style="text-align:right;font-weight:bold;">Grand Total:</td>
        <td id="grand-total">0.00</td>
      </tr>
    </tfoot>
    </table>
    """
    return table

shared_js = """
<script src="style.js"></script>
<script src="script.js"></script>
"""

shared_css = ""

# --- Generate products.html 
# <div style="display:flex; align-items:center; justify-content:space-between; margin-bottom:10px;">
#  <input type="text" id="searchInput" placeholder="Search products by name..." style="width:200px;">
# </div>
# ---
main_html = f"""
<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<title>Products - Malaiyamman Traders</title>
<link rel="stylesheet" href="https://cdn.jsdelivr.net/npm/bootstrap@5.3.2/dist/css/bootstrap.min.css">
<link rel="stylesheet" href="style.css">
<link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.5.0/css/all.min.css">
{shared_css}
</head>
<body>
{header_html}
{nav_html}
<section>
<h2>Explore Our Categories</h2>

{generate_table(df)}
</section>
{footer_html}
{shared_js}
<!-- Image Popup -->
<div id="image-popup" class="image-popup">
  <span class="close">&times;</span>
  <div class="popup-content">
    <img id="popup-img" src="" alt="Product Image">
    <div id="popup-caption" class="popup-caption"></div>
  </div>
</div>
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
<link rel="stylesheet" href="https://cdn.jsdelivr.net/npm/bootstrap@5.3.2/dist/css/bootstrap.min.css">
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
<!-- Image Popup -->
<div id="image-popup" class="image-popup">
  <span class="close">&times;</span>
  <div class="popup-content">
    <img id="popup-img" src="" alt="Product Image">
    <div id="popup-caption" class="popup-caption"></div>
  </div>
</div>
<div id="toast-container"></div>
</body>
</html>
"""
    with open(os.path.join(output_dir, filename), "w", encoding="utf-8") as f:
        f.write(cat_html)
