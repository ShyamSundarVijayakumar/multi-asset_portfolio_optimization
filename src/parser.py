import re
import pdfplumber
import pandas as pd
from pathlib import Path

"""
Broker Statement Extraction Module
----------------------------------
This module provides utilities to parse searchable PDF bank statements 
for Buy, Sell, and Dividend transactions. 

Key Features:
- Multi-language support (German/English anchors).
- Numeric cleaning for European/US decimal formats.
- Multi-line extraction (Header -> Next Line Value).
- Dynamic security name and ISIN cleaning.
"""

def clean_decimal(val):
    if not val: return "0.00"
    val = re.sub(r'[^\d.,-]', '', val)
    if not val: return "0.00"
    if ',' in val and '.' in val:
        if val.rfind('.') > val.rfind(','): val = val.replace(',', '')
        else: val = val.replace('.', '').replace(',', '.')
    elif ',' in val:
        val = val.replace(',', '.')
    try:
        return f"{float(val):.2f}"
    except:
        return "0.00"

def extract_data_from_searchable_pdf(file_path):
    data = {
        "file_original": file_path.name, 
        "date": "00.00.0000", 
        "isin": "UnknownISIN", 
        "security_name": "Unknown", 
        "type": "UnknownType", 
        "quantity": "0.00", 
        "price": "0.00", 
        "currency": "EUR", 
        "tax_withheld": "0.00", 
        "broker_fee": "0.00",
        "dividend_after_taxes": "0.00"
    }

    try:
        with pdfplumber.open(file_path) as pdf:
            # 1. PAGE SELECTION
            target_page = pdf.pages[0]
            if len(pdf.pages) > 1:
                p1_text = pdf.pages[0].extract_text() or ""
                if any(word in p1_text for word in ["Ausschüttung", "Wertpapier", "Kauf"]):
                    # If p1 is German or looks like a cover, check if p2 is cleaner English
                    target_page = pdf.pages[1]
            
            text = target_page.extract_text()
            if not text: return {**data, "status": "Error", "msg": "No text layer"}
            lines = text.split('\n')

            # 2. TRANSACTION TYPE & DATE (Search full page text)
            if re.search(r"Purchase|Kauf", text, re.I): data['type'] = "buy"
            elif re.search(r"Distribution|Dividende|Ausschüttung|Sie haben eine Dividende erhalten", text, re.I): data['type'] = "dividend"
            elif re.search(r"Sell|Verkauf", text, re.I): data['type'] = "sell"
            
            date_match = re.search(r"(\d{2}\.\d{2}\.\d{4})", text)
            if date_match: data['date'] = date_match.group(1)

            # --- 3. ISIN & SECURITY NAME  ---     
            # STEP A: Try specific German DE Logic first
            isin_match = re.search(r"\b(DE[A-Z0-9]{10,14})\b", text)
            if isin_match:
                raw_found = isin_match.group(1)
                data['isin'] = raw_found[:12] .replace('O', '0')
                
                for i, line in enumerate(lines):
                    if raw_found in line:
                        name_part = line.split(raw_found)[0].strip()
                        name_part = re.sub(r'ISIN|Wertpapier|Security|Name|Description', '', name_part, flags=re.I).strip()
                        
                        if len(name_part) < 3 and i > 0:
                            data['security_name'] = lines[i-1].strip().replace(" ", "_")
                        else:
                            data['security_name'] = name_part.replace(" ", "_")
                        break

            # STEP B: Anchor-Based Logic (Header on one line, Name/ISIN on the next)
            if data['isin'] == "UnknownISIN":
                found_by_anchor = False
                for i, line in enumerate(lines):
                    # 1. Look for the anchor header
                    if re.search(r'Wertpapier\s+ISIN|Security\s+ISIN', line, flags=re.I):
                        if i + 1 < len(lines):
                            target_line = lines[i+1].strip()
                            
                            # 2. Extract ISIN from the next line (allowing for OCR noise up to 14 chars)
                            isin_match = re.search(r'([A-Z]{2}[A-Z0-9]{10,12})', target_line)
                            
                            if isin_match:
                                full_match = isin_match.group(1)
                                
                                # Take 12 chars, force the last digit (Check Digit) to be numeric
                                raw_12 = full_match[:12]
                                data['isin'] = raw_12[:11] + raw_12[11].replace('O', '0')
                                
                                # 3. Extract Name: everything before the ISIN on that same line
                                name_part = target_line.split(full_match)[0].strip()
                                
                                # Clean Name: Remove common headers and German legal suffixes (INH, O.N.)
                                name_part = re.sub(r'ISIN|Wertpapier|Security|Name|Description|INH|O\.N\.', '', name_part, flags=re.I).strip()
                                
                                # If the name part is empty on this line, try the line above the header
                                if len(name_part) < 2 and i > 0:
                                    data['security_name'] = lines[i-1].strip().replace(" ", "_")
                                else:
                                    data['security_name'] = name_part.replace(" ", "_")
                                
                                found_by_anchor = True
                                break
                
                # --- UNIVERSAL FALLBACK ---
                # Only runs if the "Wertpapier ISIN" header logic fails
                if not found_by_anchor:
                    isin_match = re.search(r"\b([A-Z]{2}[A-Z0-9]{10,14})\b", text)
                    if isin_match:
                        full_match_str = isin_match.group(1)
                        # Standardize to 12 chars and fix last digit
                        raw_12 = full_match_str[:12]
                        data['isin'] = raw_12[:11] + raw_12[11].replace('O', '0')
                        
                        for i, line in enumerate(lines):
                            if full_match_str in line:
                                if len(line.strip()) < 25 and i > 0:
                                    data['security_name'] = lines[i-1].strip().replace(" ", "_")
                                else:
                                    # Split by keywords and take the left side
                                    parts = re.split(r'ISIN|Wertpapier|Security|Description', line, flags=re.I)
                                    name_raw = parts[0].strip()
                                    name_clean = re.sub(r'INH|O\.N\.', '', name_raw, flags=re.I).strip()
                                    data['security_name'] = name_clean.replace(" ", "_")
                                break
            
            # Clean up any trailing dots or noise from security name
            if data['security_name'] != "Unknown":
                data['security_name'] = re.sub(r'INH|O\.N\.', '', data['security_name'], flags=re.I).strip('_').strip('.')

            # 4. QUANTITY & PRICE
            for i, line in enumerate(lines):
                if re.search(r"Nominal|Anzahl|Quantity|Units", line, re.I):
                    if i + 1 < len(lines):
                        next_line = lines[i+1]
                        nums = re.findall(r"[\d,.]+", next_line)
                        nums = [n for n in nums if len(n) > 1 and not re.match(r'^\d{2}\.\d{2}\.\d{4}$', n)]
                        if len(nums) >= 2:
                            data['quantity'] = clean_decimal(nums[0])
                            data['price'] = clean_decimal(nums[1])
                    break

            # 5. TAXES & FEES
            total_tax = 0.0
            for line in lines:
                if re.search(r"Broker service fee|Commission", line, re.I):
                    f_match = re.search(r"([\d,.]+)", line)
                    if f_match: data['broker_fee'] = clean_decimal(f_match.group(1)) 
                
                if re.search(r"Capital gains tax|Solidarity surcharge|German withholding tax|Quellensteuer|Kirchensteuer", line, re.I):
                    t_nums = re.findall(r"[\d,.]+", line)
                    if t_nums:
                        total_tax += float(clean_decimal(t_nums[-1]))
            
            data['tax_withheld'] = f"{total_tax:.2f}"

            # Only extract for dividends; keep as 0.00 for buy/sell
            # 6. FINAL DATE, CURRENCY & NET DIVIDEND
            for i, line in enumerate(lines):
                if re.search(r"Value date|Amount credited|Betrag zu Ihren Gunsten", line, re.I):
                    if i + 1 < len(lines):
                        next_line = lines[i+1].strip()
                        net_match = re.search(r"(\d{2}\.\d{2}\.\d{4})\s+([\d,.]+)\s+([A-Z]{3})", next_line)
                        if net_match:
                            data['date'] = net_match.group(1)
                            data['currency'] = net_match.group(3)
                            
                            # Only save the money amount to 'dividend_after_taxes' if it's a dividend
                            if data['type'] == "dividend":
                                data['dividend_after_taxes'] = clean_decimal(net_match.group(2))
                            else:
                                data['dividend_after_taxes'] = "0.00"
                            break


            # Name Clean up (Removes INH, O.N. and leading/trailing underscores/dots)
            if data['security_name'] != "Unknown":
                data['security_name'] = re.sub(r'INH|O\.N\.', '', data['security_name'], flags=re.I).strip('_').strip('.')

            return data

    except Exception as e:
        return {**data, "status": "Error", "msg": str(e)}


# for source b
def extract_source_b(file_path, text):
    lines = text.split('\n')
    data = {
        "file_original": file_path.name,
        "date": "00.00.0000",
        "isin": "Unknown",
        "security_name": "Unknown",
        "type": "UnknownType",
        "quantity": "0.00",
        "price": "0.00",
        "currency": "EUR",
        "tax_withheld": "0.00",
        "broker_fee": "0.00",
        "dividend_after_taxes": "0.00"
    }

    # 1. Extract Date and ISIN from filename (Format: YYYY-MM-DD-Type-ISIN)
    file_match = re.search(r"(\d{4}-\d{2}-\d{2}).*?([A-Z]{2}[A-Z0-9]{10})", file_path.name)
    if file_match:
        data['date'] = file_match.group(1)
        data['isin'] = file_match.group(2)

    for line in lines:
        line = line.strip()
        if not line:
            continue

        # --- 2. EXTRACT SECURITY NAME (Corporate Action/Dividend Layout) ---
        if "Entitled security" in line:
            name_match = re.search(r"Entitled security\s+(.*)", line, re.I)
            if name_match:
                data['security_name'] = name_match.group(1).strip().replace(" ", "_")

        # --- LAYOUT 1: Buy/Sell  ---
        if any(kw in line for kw in ["Buy", "Sell"]):
            data['type'] = "buy" if "Buy" in line else "sell"
            
            # 1. Extract Security Name
            name_regex = re.search(r"(?:Buy|Sell)\s+(.*?)\s+\d", line)
            if name_regex and data['security_name'] == "Unknown":
                data['security_name'] = name_regex.group(1).strip().replace(" ", "_")
            
            # 2. Extract Numbers (Quantity and Price)
            nums = re.findall(r"\d+[\.,]\d+", line)
            if len(nums) >= 2:
                data['quantity'] = clean_decimal(nums[0])
                data['price'] = clean_decimal(nums[1])
            
            # 3. Extract Currency
            # This looks for 3 uppercase letters that follow a space and a digit
            # OR are at the very end of the line, avoiding words in the name.
            curr_match = re.search(r"\d\s+([A-Z]{3})(?:\s|$)", line)
            if curr_match: 
                data['currency'] = curr_match.group(1)

        # --- 4. LAYOUT 2: Dividends ---
        if "Credit" in line and "EUR" in line:
            data['type'] = "dividend"
            # Pattern: Date Credit [Price] EUR [Qty] [Total] EUR
            nums = re.findall(r"[\d,.]+", line)
            if len(nums) >= 4:
                # We prioritize the structure: Price (idx 1), Quantity (idx 2)
                data['price'] = clean_decimal(nums[1])    
                data['quantity'] = clean_decimal(nums[2]) 

        # --- 5. GLOBAL FIELDS (Fees and Taxes) ---
        if "Taxes" in line:
            tax_match = re.search(r"(-?[\d,.]+)", line)
            if tax_match:
                data['tax_withheld'] = f"{abs(float(clean_decimal(tax_match.group(1)))):.2f}"

        if "Total" in line and data['type'] == "dividend":
            total_match = re.search(r"([\d,.]+)", line)
            if total_match:
                data['dividend_after_taxes'] = clean_decimal(total_match.group(1))

        if "Order fees" in line:
            fee_match = re.search(r"[\d,.]+", line)
            if fee_match: 
                data['broker_fee'] = clean_decimal(fee_match.group(0))

    return data

# for source c
def extract_source_c_csv(file_path):
    df = pd.read_csv(file_path)
    
    # Filter for relevant actions
    relevant_actions = ["Market buy", "Market sell", "Dividend", "Distribution"]
    df = df[df['Action'].str.contains('|'.join(relevant_actions), case=False, na=False)].copy()
    
    results = []
    for _, row in df.iterrows():
        action = str(row['Action']).lower()
        
        # Normalize type
        trans_type = "dividend" if "div" in action or "dist" in action else action.replace("market ", "")
        
        # --- Safely Extract Exchange Rate ---
        # Default to 1.0 to prevent division by zero errors
        try:
            ex_rate_str = str(row.get('Exchange rate', '1')).replace(',', '.')
            exchange_rate = float(re.sub(r'[^\d.]', '', ex_rate_str))
            if exchange_rate == 0: exchange_rate = 1.0
        except ValueError:
            exchange_rate = 1.0

        # --- Extract numeric values using your clean_decimal function ---
        raw_price = float(clean_decimal(str(row.get('Price / share', '0'))))
        raw_tax   = float(clean_decimal(str(row.get('Withholding tax', '0'))))
        raw_fee   = float(clean_decimal(str(row.get('Currency conversion fee', '0'))))
        raw_total = float(clean_decimal(str(row.get('Total', '0'))))
        
        # --- Currency Conversions (USD to EUR) ---
        
        # 1. Price Conversion
        price_curr = str(row.get('Currency (Price / share)', '')).strip().upper()
        result_curr = str(row.get('Currency (Result)', '')).strip().upper()
        
        # If Price is USD and Result is EUR, or it's a USD Dividend
        if price_curr == 'USD' and (result_curr == 'EUR' or trans_type == 'dividend'):
            raw_price = raw_price / exchange_rate
            
        # 2. Tax Conversion
        tax_curr = str(row.get('Currency (Withholding tax)', '')).strip().upper()
        if tax_curr == 'USD':
            raw_tax = raw_tax / exchange_rate
            
        # 3. Broker Fee Conversion
        fee_curr = str(row.get('Currency (Currency conversion fee)', '')).strip().upper()
        if fee_curr == 'USD':
            raw_fee = raw_fee / exchange_rate

        # --- Build Output Dictionary ---
        data = {
            "file_original": file_path.name,
            "date": str(row['Time']).split()[0],
            "isin": row.get('ISIN', 'Unknown'),
            "security_name": str(row.get('Name', 'Unknown')).replace(" ", "_"),
            "type": trans_type,
            "quantity": clean_decimal(str(row.get('No. of shares', '0'))),
            "price": f"{raw_price:.2f}",
            "currency": "EUR", # standardize to EUR after conversions
            "tax_withheld": f"{raw_tax:.2f}",
            "broker_fee": f"{raw_fee:.2f}",
            "dividend_after_taxes": "0.00"
        }
        
        # --- Dividend Calculation ---
        # Logic: Total - Withholding tax
        if trans_type == "dividend":
            div_after_taxes = raw_total - raw_tax
            data["dividend_after_taxes"] = f"{div_after_taxes:.2f}"
            
        results.append(data)
        
    return results