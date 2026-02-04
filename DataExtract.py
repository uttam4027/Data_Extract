"""
Jewelry Data Processor - Complete Solution
===========================================
Upload Excel → Extract URLs → Download Images → Process to Square → Insert in Excel

Author: Claude
Date: 2026-02-04
Version: 2.0 (Complete image processing pipeline)
"""

import streamlit as st
import pandas as pd
import re
import win32com.client
import pythoncom
import time
from pathlib import Path
from io import BytesIO
import tempfile
import os
import threading
from queue import Queue
import subprocess
import requests
import cv2
import numpy as np
from PIL import Image
import random

# ======================
# PAGE CONFIGURATION
# ======================
st.set_page_config(
    page_title="Jewelry Data Processor - Complete",
    page_icon="💎",
    layout="wide"
)

# ======================
# VBA MACRO CODES
# ======================
VBA_ExtractR_CODE = '''
Option Explicit

Sub ExtractURLsFromColumnA_Auto()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim currentRow As Long
    Dim url As String
    Dim ExtractdData As String
    Dim startTime As Double
    Dim successCount As Long
    Dim failCount As Long
    
    On Error GoTo ErrorHandler
    
    Set ws = ThisWorkbook.ActiveSheet
    successCount = 0
    failCount = 0
    startTime = Timer
    
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False
    Application.DisplayStatusBar = True
    
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    
    If ws.Cells(1, 1).Value = "" Then ws.Cells(1, 1).Value = "URL"
    If ws.Cells(1, 3).Value = "" Then ws.Cells(1, 3).Value = "Extractd Data"
    
    For currentRow = 2 To lastRow
        url = Trim(ws.Cells(currentRow, 1).Value)
        
        Application.StatusBar = "Processing " & (currentRow - 1) & " of " & (lastRow - 1) & ": " & Left(url, 50)
        
        If url <> "" Then
            ExtractdData = ExtractURL(url)
            
            If Left(ExtractdData, 6) = "ERROR:" Then
                ws.Cells(currentRow, 3).Value = ExtractdData
                ws.Cells(currentRow, 3).Interior.Color = RGB(255, 200, 200)
                failCount = failCount + 1
            Else
                ws.Cells(currentRow, 3).Value = ExtractdData
                ws.Cells(currentRow, 3).Interior.ColorIndex = xlNone
                successCount = successCount + 1
            End If
            
            If currentRow Mod 5 = 0 Then
                DoEvents
                Application.Wait Now + TimeValue("00:00:01")
            End If
        End If
    Next currentRow
    
    ws.Range("Z1").Value = "SUCCESS|" & successCount & "|" & failCount & "|" & Format((Timer - startTime) / 60, "0.0")
    
    GoTo CleanUp
    
ErrorHandler:
    ws.Range("Z1").Value = "ERROR|" & Err.Description
    
CleanUp:
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    Application.StatusBar = False
    ws.Columns("A:C").AutoFit
End Sub

Function ExtractURL(url As String) As String
    Dim http As Object
    Dim htmlDoc As Object
    Dim htmlBody As Object
    Dim textContent As String
    Dim timeout As Long
    
    On Error GoTo ErrorHandler
    
    If Not (InStr(1, url, "http://", vbTextCompare) > 0 Or InStr(1, url, "https://", vbTextCompare) > 0) Then
        ExtractURL = "ERROR: Invalid URL format"
        Exit Function
    End If
    
    Set http = CreateObject("MSXML2.ServerXMLHTTP.6.0")
    timeout = 30000
    http.setTimeouts timeout, timeout, timeout, timeout
    
    http.Open "GET", url, False
    http.setRequestHeader "User-Agent", "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36"
    http.Send
    
    If http.Status = 200 Then
        Set htmlDoc = CreateObject("HTMLFile")
        htmlDoc.body.innerHTML = http.responseText
        
        Set htmlBody = htmlDoc.body
        If Not htmlBody Is Nothing Then
            textContent = CleanText(htmlBody.innerText)
            
            If Len(textContent) > 32000 Then
                textContent = Left(textContent, 32000) & "... [TRUNCATED]"
            End If
            
            ExtractURL = textContent
        Else
            ExtractURL = "ERROR: No body content found"
        End If
    Else
        ExtractURL = "ERROR: HTTP " & http.Status & " - " & http.statusText
    End If
    
    Set http = Nothing
    Set htmlDoc = Nothing
    Set htmlBody = Nothing
    
    Exit Function
    
ErrorHandler:
    ExtractURL = "ERROR: " & Err.Description
    Set http = Nothing
    Set htmlDoc = Nothing
    Set htmlBody = Nothing
End Function

Function CleanText(rawText As String) As String
    Dim result As String
    result = rawText
    
    result = Trim(result)
    
    Do While InStr(result, "  ") > 0
        result = Replace(result, "  ", " ")
    Loop
    
    Do While InStr(result, vbCrLf & vbCrLf & vbCrLf) > 0
        result = Replace(result, vbCrLf & vbCrLf & vbCrLf, vbCrLf & vbCrLf)
    Loop
    
    CleanText = result
End Function
'''

VBA_IMAGE_INSERT_CODE = '''
Sub InsertImagesFromPaths_Auto()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim imagePath As String
    Dim shp As Shape
    Dim cellWidth As Double
    Dim cellHeight As Double
    Dim imgRatio As Double
    Dim imageColumn As Long
    Dim targetColumn As Long
    Dim missingCount As Long
    Dim successCount As Long
    
    On Error GoTo ErrorHandler
    
    Set ws = ThisWorkbook.ActiveSheet
    
    ' Find image path column (should be last column)
    lastRow = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
    imageColumn = lastRow
    targetColumn = imageColumn + 1
    
    ' Add header for image column
    ws.Cells(1, targetColumn).Value = "Product Image"
    
    ' Find last row with data
    lastRow = ws.Cells(ws.Rows.Count, imageColumn).End(xlUp).Row
    
    ' Set column width
    ws.Columns(targetColumn).ColumnWidth = 20
    
    ' Turn off screen updating
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    
    ' Initialize counters
    missingCount = 0
    successCount = 0
    
    ' Loop through each row
    For i = 2 To lastRow
        imagePath = Trim(ws.Cells(i, imageColumn).Value)
        
        If imagePath <> "" And Dir(imagePath) <> "" Then
            ' Set row height
            ws.Rows(i).RowHeight = 100
            
            ' Get cell dimensions
            cellWidth = ws.Cells(i, targetColumn).Width
            cellHeight = ws.Cells(i, targetColumn).Height
            
            ' Insert and embed picture
            Set shp = ws.Shapes.AddPicture( _
                Filename:=imagePath, _
                LinkToFile:=msoFalse, _
                SaveWithDocument:=msoTrue, _
                Left:=ws.Cells(i, targetColumn).Left, _
                Top:=ws.Cells(i, targetColumn).Top, _
                Width:=-1, _
                Height:=-1)
            
            ' Calculate aspect ratio and resize
            imgRatio = shp.Width / shp.Height
            If imgRatio > (cellWidth / cellHeight) Then
                shp.Width = cellWidth - 4
                shp.Height = shp.Width / imgRatio
            Else
                shp.Height = cellHeight - 4
                shp.Width = shp.Height * imgRatio
            End If
            
            ' Center image in cell
            shp.Top = ws.Cells(i, targetColumn).Top + (cellHeight - shp.Height) / 2
            shp.Left = ws.Cells(i, targetColumn).Left + (cellWidth - shp.Width) / 2
            
            shp.Placement = xlMoveAndSize
            shp.LockAspectRatio = msoTrue
            shp.Name = "IMG_Row" & i
            
            successCount = successCount + 1
        Else
            missingCount = missingCount + 1
        End If
    Next i
    
    ws.Range("Z2").Value = "SUCCESS|" & successCount & "|" & missingCount
    
    GoTo CleanUp
    
ErrorHandler:
    ws.Range("Z2").Value = "ERROR|" & Err.Description
    
CleanUp:
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
End Sub
'''

# ======================
# IMAGE PROCESSING
# ======================
class ImageProcessor:
    def __init__(self, output_folder):
        self.output_folder = output_folder
        os.makedirs(output_folder, exist_ok=True)
        self.session = requests.Session()
        self.setup_session()
    
    def setup_session(self):
        """Setup session with headers"""
        headers = {
            'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36',
            'Accept': 'image/webp,image/apng,image/*,*/*;q=0.8',
            'Accept-Language': 'en-US,en;q=0.9',
            'Referer': 'https://www.google.com/'
        }
        self.session.headers.update(headers)
    
    def download_image(self, url, save_path):
        """Download image with multiple fallback methods"""
        # Method 1: Session with headers
        try:
            time.sleep(random.uniform(0.3, 0.8))
            response = self.session.get(url, timeout=30, allow_redirects=True)
            response.raise_for_status()
            
            # Verify content is actually an image
            content_type = response.headers.get('content-type', '')
            if 'image' not in content_type.lower() and len(response.content) < 1000:
                return False
            
            with open(save_path, 'wb') as f:
                f.write(response.content)
            
            # Verify file was written
            return os.path.exists(save_path) and os.path.getsize(save_path) > 0
            
        except Exception as e1:
            # Method 2: Simple requests without session
            try:
                response = requests.get(
                    url, 
                    timeout=30, 
                    headers={
                        'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36',
                        'Accept': 'image/*'
                    }
                )
                response.raise_for_status()
                
                with open(save_path, 'wb') as f:
                    f.write(response.content)
                
                return os.path.exists(save_path) and os.path.getsize(save_path) > 0
                
            except Exception as e2:
                # Method 3: Try with urllib as last resort
                try:
                    import urllib.request
                    req = urllib.request.Request(
                        url,
                        headers={
                            'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36'
                        }
                    )
                    with urllib.request.urlopen(req, timeout=30) as response:
                        with open(save_path, 'wb') as f:
                            f.write(response.read())
                    
                    return os.path.exists(save_path) and os.path.getsize(save_path) > 0
                except:
                    return False
    
    def make_square_image(self, image_path, offset=20):
        """Convert image to square format with detected background"""
        try:
            img = cv2.imread(image_path)
            if img is None:
                return False
            
            gray = cv2.cvtColor(img, cv2.COLOR_BGR2GRAY)
            h, w = img.shape[:2]
            
            # Detect background color from corners
            corner_samples = [
                img[0:10, 0:10],
                img[0:10, w-10:w],
                img[h-10:h, 0:10],
                img[h-10:h, w-10:w]
            ]
            all_samples = np.vstack([s.reshape(-1, 3) for s in corner_samples])
            background_color = np.mean(all_samples, axis=0).astype(np.uint8)
            
            # Create mask
            _, mask = cv2.threshold(gray, 240, 255, cv2.THRESH_BINARY)
            mask_inv = cv2.bitwise_not(mask)
            kernel = np.ones((3,3), np.uint8)
            mask_inv = cv2.morphologyEx(mask_inv, cv2.MORPH_CLOSE, kernel)
            mask_inv = cv2.morphologyEx(mask_inv, cv2.MORPH_OPEN, kernel)
            
            # Find contours
            contours, _ = cv2.findContours(mask_inv, cv2.RETR_EXTERNAL, cv2.CHAIN_APPROX_SIMPLE)
            
            if contours:
                filtered_contours = [c for c in contours if cv2.contourArea(c) > 100]
                if not filtered_contours:
                    return False
                
                # Get bounding box
                all_points = np.vstack(filtered_contours)
                x, y, w_box, h_box = cv2.boundingRect(all_points)
                
                # Add extra 10px area
                extra = 10
                x_extra = max(0, x - extra)
                y_extra = max(0, y - extra)
                w_extra = min(img.shape[1] - x_extra, w_box + 2 * extra)
                h_extra = min(img.shape[0] - y_extra, h_box + 2 * extra)
                
                # Add offset
                x_offset = max(0, x_extra - offset)
                y_offset = max(0, y_extra - offset)
                w_offset = min(img.shape[1] - x_offset, w_extra + 2 * offset)
                h_offset = min(img.shape[0] - y_offset, h_extra + 2 * offset)
                
                # Crop product
                cropped = img[y_offset:y_offset + h_offset, x_offset:x_offset + w_offset]
                
                # Create square canvas
                max_dim = max(w_offset, h_offset)
                square_img = np.full((max_dim, max_dim, 3), background_color, dtype=np.uint8)
                
                # Center the product
                start_x = (max_dim - w_offset) // 2
                start_y = (max_dim - h_offset) // 2
                square_img[start_y:start_y + h_offset, start_x:start_x + w_offset] = cropped
                
                # Save result
                cv2.imwrite(image_path, square_img)
                return True
            return False
        except:
            return False
    
    def process_row(self, sku, image_url, progress_callback=None):
        """Download and process a single image"""
        try:
            # Convert to string and check if valid
            image_url = str(image_url).strip()
            
            if not image_url or image_url.lower() in ['nan', 'none', '']:
                if progress_callback:
                    progress_callback(f"Skipping {sku}: No URL")
                return None, "No URL"
            
            # Check if URL is valid
            if not image_url.startswith(('http://', 'https://')):
                if progress_callback:
                    progress_callback(f"Skipping {sku}: Invalid URL format")
                return None, "Invalid URL format"
            
            # Download image
            filename = f"{sku}.jpg"
            temp_path = os.path.join(self.output_folder, filename)
            
            if progress_callback:
                progress_callback(f"Downloading {sku} from {image_url[:50]}...")
            
            if not self.download_image(image_url, temp_path):
                if progress_callback:
                    progress_callback(f"Failed to download {sku}")
                return None, "Download failed"
            
            # Verify file was created and has content
            if not os.path.exists(temp_path) or os.path.getsize(temp_path) == 0:
                if progress_callback:
                    progress_callback(f"Download verification failed for {sku}")
                return None, "File not created or empty"
            
            # Process to square
            if progress_callback:
                progress_callback(f"Processing {sku} to square...")
            
            if not self.make_square_image(temp_path, offset=20):
                # Return original path even if square conversion failed
                if progress_callback:
                    progress_callback(f"Square conversion partial for {sku}")
                return temp_path, "Using original (square conversion partial)"
            
            if progress_callback:
                progress_callback(f"✓ Completed {sku}")
            
            return temp_path, "Success"
        
        except Exception as e:
            if progress_callback:
                progress_callback(f"Error processing {sku}: {str(e)}")
            return None, str(e)

# ======================
# HELPER FUNCTIONS
# ======================
def clean_text(text):
    """Clean text by removing special characters"""
    if pd.isna(text):
        return ""
    text = str(text)
    text = text.replace("_x000D_", "\n")
    text = text.replace("\r", "\n")
    return text

def extract_prices(line):
    """Extract listing and discounted prices"""
    prices = re.findall(r'\$([\d,]+\.\d{2})', line)
    prices = [float(p.replace(",", "")) for p in prices]
    if len(prices) >= 2:
        return max(prices), min(prices)
    elif len(prices) == 1:
        return prices[0], None
    return None, None

def parse_top_block(text):
    """Parse top block to extract SKU, status, description, prices"""
    result = {
        "SKU_Actual": None,
        "Status": None,
        "Desc_Text": None,
        "Listing_Price": None,
        "Discounted_Price": None
    }
    
    if "Item #:" not in text:
        return None
    
    text = text[text.index("Item #:"):]
    lines = [l.strip() for l in text.split("\n") if l.strip()]
    
    sku_match = re.search(r'Item\s*#:\s*(\d+)', text)
    if sku_match:
        result["SKU_Actual"] = sku_match.group(1)
    
    price_idx = None
    for i, l in enumerate(lines):
        if "price" in l.lower() and "$" in l:
            listing, discounted = extract_prices(l)
            result["Listing_Price"] = listing
            result["Discounted_Price"] = discounted
            price_idx = i
            break
    
    if price_idx is None:
        return result
    
    middle_lines = lines[1:price_idx]
    
    if len(middle_lines) == 2:
        result["Status"] = middle_lines[0]
        result["Desc_Text"] = middle_lines[1]
    elif len(middle_lines) == 1:
        result["Desc_Text"] = middle_lines[0]
    
    return result

def extract_details_block(text):
    """Extract details block with stone information"""
    start = re.search(r'Details\s*Stone\(s\)', text, re.I)
    end = re.search(r'Shipping and Returns', text, re.I)
    
    if not start or not end:
        return []
    
    block = text[start.end():end.start()]
    return [l.strip() for l in block.split("\n") if l.strip()]

def process_Extractd_data(url_df, map_df, progress_bar, status_text):
    """Process Extractd data into structured format"""
    mapping = dict(zip(map_df["Online"], map_df["Main"]))
    results = []
    total = len(url_df)
    
    for idx, row in url_df.iterrows():
        progress = (idx + 1) / total
        progress_bar.progress(progress)
        status_text.text(f"Processing record {idx + 1} of {total}...")
        
        raw = clean_text(row.get("Extractd Data", ""))
        
        if raw.startswith("ERROR:") or not raw:
            continue
        
        base = parse_top_block(raw)
        
        if base is None:
            continue
        
        for col in mapping.values():
            base[col] = None
        
        details = extract_details_block(raw)
        for line in details:
            for online, main in mapping.items():
                if line.startswith(online):
                    base[main] = line.replace(online, "").strip()
                    break
        
        # Add URL and Image_src columns
        base['URL'] = row.get('URL', '')
        base['Image_src'] = row.get('Image_src', '')
        
        results.append(base)
    
    progress_bar.progress(1.0)
    status_text.text(f"✓ Processed {len(results)} valid records")
    
    return pd.DataFrame(results)

# ======================
# EXCEL OPERATIONS (THREADED)
# ======================
def excel_Extractr_worker(file_path, result_queue):
    """Worker function for Web extracting"""
    try:
        pythoncom.CoInitialize()
        excel = None
        workbook = None
        
        try:
            result_queue.put(("status", "Starting Excel..."))
            excel = win32com.client.Dispatch("Excel.Application")
            
            excel.Visible = False
            excel.DisplayAlerts = False
            excel.ScreenUpdating = False
            excel.Interactive = False
            excel.EnableEvents = False
            excel.AskToUpdateLinks = False
            excel.AlertBeforeOverwriting = False
            
            result_queue.put(("progress", 0.1))
            result_queue.put(("status", "Opening workbook..."))
            
            # Convert to absolute path
            abs_path = os.path.abspath(file_path)
            
            workbook = excel.Workbooks.Open(
                abs_path,
                UpdateLinks=0,
                ReadOnly=False,
                Password='',
                WriteResPassword='',
                IgnoreReadOnlyRecommended=True,
                Notify=False,
                AddToMru=False,
                Local=False,
                CorruptLoad=0
            )
            
            if workbook is None:
                result_queue.put(("error", "Failed to open workbook"))
                return
            
            result_queue.put(("progress", 0.2))
            result_queue.put(("status", "Injecting VBA Extractr..."))
            
            # Access VBA project
            try:
                vb_project = workbook.VBProject
            except Exception as e:
                result_queue.put(("error", f"Cannot access VBA project. Please enable 'Trust access to VBA project object model' in Excel: {str(e)}"))
                return
            
            for component in vb_project.VBComponents:
                if component.Name == "ExtractrModule":
                    vb_project.VBComponents.Remove(component)
            
            module = vb_project.VBComponents.Add(1)
            module.Name = "ExtractrModule"
            module.CodeModule.AddFromString(VBA_ExtractR_CODE)
            
            result_queue.put(("progress", 0.3))
            result_queue.put(("status", "Running web Extractr..."))
            
            excel.Run("ExtractrModule.ExtractURLsFromColumnA_Auto")
            time.sleep(2)
            
            result_queue.put(("progress", 0.8))
            result_queue.put(("status", "Reading results..."))
            
            ws = workbook.ActiveSheet
            result = ws.Range("Z1").Value
            
            if result and result.startswith("SUCCESS"):
                parts = result.split("|")
                ws.Range("Z1").Value = ""
                workbook.Save()
                
                used_range = ws.UsedRange
                data = used_range.Value
                df = pd.DataFrame(data[1:], columns=data[0])
                
                result_queue.put(("progress", 1.0))
                result_queue.put(("success", df, f"Success: {parts[1]} URLs, Failed: {parts[2]} URLs"))
            else:
                result_queue.put(("error", f"Extracting failed: {result}"))
                
        except Exception as e:
            result_queue.put(("error", f"Error: {str(e)}"))
            
        finally:
            if workbook:
                try:
                    workbook.Close(SaveChanges=True)
                except:
                    pass
            if excel:
                try:
                    excel.Quit()
                except:
                    pass
    finally:
        pythoncom.CoUninitialize()

def excel_image_insert_worker(file_path, result_queue):
    """Worker function for inserting images"""
    try:
        pythoncom.CoInitialize()
        excel = None
        workbook = None
        
        try:
            result_queue.put(("status", "Opening workbook for image insertion..."))
            excel = win32com.client.Dispatch("Excel.Application")
            
            excel.Visible = False
            excel.DisplayAlerts = False
            excel.ScreenUpdating = False
            excel.Interactive = False
            excel.EnableEvents = False
            
            abs_path = os.path.abspath(file_path)
            
            workbook = excel.Workbooks.Open(
                abs_path,
                UpdateLinks=0,
                ReadOnly=False,
                Password='',
                IgnoreReadOnlyRecommended=True,
                Notify=False,
                AddToMru=False
            )
            
            if workbook is None:
                result_queue.put(("error", "Failed to open workbook"))
                return
            
            result_queue.put(("status", "Injecting image insertion macro..."))
            
            try:
                vb_project = workbook.VBProject
            except Exception as e:
                result_queue.put(("error", f"Cannot access VBA project: {str(e)}"))
                return
            
            for component in vb_project.VBComponents:
                if component.Name == "ImageModule":
                    vb_project.VBComponents.Remove(component)
            
            module = vb_project.VBComponents.Add(1)
            module.Name = "ImageModule"
            module.CodeModule.AddFromString(VBA_IMAGE_INSERT_CODE)
            
            result_queue.put(("status", "Inserting images into Excel..."))
            
            excel.Run("ImageModule.InsertImagesFromPaths_Auto")
            time.sleep(2)
            
            ws = workbook.ActiveSheet
            result = ws.Range("Z2").Value
            
            if result and result.startswith("SUCCESS"):
                parts = result.split("|")
                ws.Range("Z2").Value = ""
                workbook.Save()
                result_queue.put(("success", f"Inserted {parts[1]} images, {parts[2]} missing"))
            else:
                result_queue.put(("error", f"Image insertion failed: {result}"))
                
        except Exception as e:
            result_queue.put(("error", f"Error: {str(e)}"))
            
        finally:
            if workbook:
                try:
                    workbook.Close(SaveChanges=True)
                except:
                    pass
            if excel:
                try:
                    excel.Quit()
                except:
                    pass
    finally:
        pythoncom.CoUninitialize()

def run_threaded_operation(worker_func, file_path, progress_bar, status_text):
    """Run Excel operation in thread"""
    result_queue = Queue()
    worker_thread = threading.Thread(target=worker_func, args=(file_path, result_queue))
    worker_thread.start()
    
    while worker_thread.is_alive() or not result_queue.empty():
        try:
            msg = result_queue.get(timeout=0.1)
            msg_type = msg[0]
            
            if msg_type == "status":
                status_text.text(msg[1])
            elif msg_type == "progress":
                progress_bar.progress(msg[1])
            elif msg_type == "success":
                if len(msg) > 2:
                    return msg[1], True, msg[2]
                else:
                    return None, True, msg[1]
            elif msg_type == "error":
                return None, False, msg[1]
        except:
            continue
    
    worker_thread.join()
    return None, False, "Unknown error"

# ======================
# STREAMLIT UI
# ======================
def main():
    st.title("💎 Jewelry Data Processor - Complete Solution")
    st.markdown("### Upload → Extract → Download Images → Process to Square → Insert in Excel")
    st.markdown("---")
    
    # Sidebar
    with st.sidebar:
        st.header("⚙️ Settings")
        st.markdown("---")
        
        # VBA Setup Check
        with st.expander("⚠️ VBA Setup Required", expanded=True):
            st.markdown("""
            **One-time Excel setup:**
            1. Open Excel
            2. File → Options → Trust Center
            3. Trust Center Settings
            4. Macro Settings
            5. ✅ Enable: **"Trust access to VBA project object model"**
            6. Click OK
            
            Without this, macro injection will fail.
            """)
        
        st.markdown("---")
        
        enable_images = st.checkbox("📸 Download & Process Images", value=True, 
                                    help="Download images from Image_src column and convert to square format")
        
        if enable_images:
            st.caption("Images will be:")
            st.caption("• Downloaded from Image_src URLs")
            st.caption("• Converted to square format")
            st.caption("• Inserted in final Excel")
        
        st.markdown("---")
        
        if st.button("🧹 Force Close Excel", help="Click if Excel is stuck"):
            try:
                subprocess.run(['taskkill', '/F', '/IM', 'EXCEL.EXE'], 
                              capture_output=True, timeout=5)
                st.success("✓ Excel processes cleaned")
                time.sleep(1)
                st.rerun()
            except:
                st.warning("No Excel processes found")
        
        st.markdown("---")
        st.caption("Excel runs in background")
        st.caption("No windows will open")
    
    # File upload
    uploaded_file = st.file_uploader(
        "Upload Excel file (Columns: A=URL, B=Image_src, + Mapping sheet)",
        type=['xlsx', 'xls'],
        help="Required: URL in column A, Image_src in column B, and 'Mapping' sheet"
    )
    
    if uploaded_file is not None:
        col1, col2 = st.columns(2)
        with col1:
            st.info(f"📄 **File:** {uploaded_file.name}")
        with col2:
            st.info(f"📊 **Size:** {uploaded_file.size / 1024:.2f} KB")
        
        temp_dir = tempfile.mkdtemp()
        temp_path = os.path.join(temp_dir, uploaded_file.name)
        
        with open(temp_path, 'wb') as f:
            f.write(uploaded_file.getbuffer())
        
        st.success("✓ File uploaded successfully!")
        
        # Preview
        with st.expander("📋 Preview Data", expanded=False):
            try:
                preview_df = pd.read_excel(temp_path, nrows=5)
                st.dataframe(preview_df, use_container_width=True)
            except Exception as e:
                st.error(f"Preview error: {e}")
        
        st.markdown("---")
        st.info("ℹ️ Excel runs completely in background - no windows will open")
        
        if st.button("🚀 Start Complete Processing", type="primary", use_container_width=True):
            
            # STEP 1: Web extracting
            st.markdown("### 📡 Step 1: Web extracting")
            progress_bar1 = st.progress(0)
            status_text1 = st.empty()
            
            with st.spinner("Extracting URLs..."):
                url_df, success, message = run_threaded_operation(
                    excel_Extractr_worker, temp_path, progress_bar1, status_text1
                )
            
            if not success:
                st.error(f"❌ Extracting failed: {message}")
                return
            
            st.success(f"✓ {message}")
            
            # STEP 2: Download & Process Images (if enabled)
            image_paths = []
            if enable_images:
                st.markdown("### 📸 Step 2: Download & Process Images")
                
                # Show debug info about Image_src column
                with st.expander("🔍 Debug: Image_src URLs Found", expanded=False):
                    image_src_count = url_df['Image_src'].notna().sum()
                    st.write(f"**Total rows:** {len(url_df)}")
                    st.write(f"**Rows with Image_src:** {image_src_count}")
                    st.write(f"**Empty Image_src:** {len(url_df) - image_src_count}")
                    
                    # Show first few Image_src URLs
                    if image_src_count > 0:
                        st.write("**Sample Image_src URLs:**")
                        sample_urls = url_df[url_df['Image_src'].notna()]['Image_src'].head(5)
                        for i, url in enumerate(sample_urls, 1):
                            st.text(f"{i}. {url}")
                
                images_folder = os.path.join(temp_dir, "processed_images")
                processor = ImageProcessor(images_folder)
                
                progress_bar2 = st.progress(0)
                status_text2 = st.empty()
                
                success_count = 0
                fail_count = 0
                
                total_rows = len(url_df)
                
                # Create a mapping of Image_src to image paths
                image_src_to_path = {}
                
                for idx, row in url_df.iterrows():
                    # Use row index as SKU for filename
                    sku = f"product_{idx + 1}"
                    image_url = row.get('Image_src', '')
                    
                    progress = (idx + 1) / total_rows
                    progress_bar2.progress(progress)
                    
                    if pd.isna(image_url) or str(image_url).strip() == '' or str(image_url).lower() == 'nan':
                        status_text2.text(f"Skipping row {idx + 1}: No Image_src URL")
                        image_paths.append('')
                        fail_count += 1
                        continue
                    
                    path, status = processor.process_row(
                        sku, image_url,
                        lambda msg: status_text2.text(msg)
                    )
                    
                    if path and os.path.exists(path):
                        image_paths.append(path)
                        image_src_to_path[image_url] = path
                        success_count += 1
                    else:
                        image_paths.append('')
                        fail_count += 1
                
                progress_bar2.progress(1.0)
                status_text2.text(f"✓ Downloaded: {success_count}, Failed: {fail_count}")
                
                st.success(f"✓ Images processed: {success_count} successful, {fail_count} failed")
                
                # Add image paths to DataFrame
                url_df['Processed_Image_Path'] = image_paths
                url_df['Image_src_to_path'] = url_df['Image_src'].map(image_src_to_path).fillna('')
            
            # STEP 3: Process Data
            st.markdown("### 🔄 Step 3: Processing Data")
            progress_bar3 = st.progress(0)
            status_text3 = st.empty()
            
            try:
                map_df = pd.read_excel(temp_path, sheet_name="Mapping")
                st.info(f"✓ Loaded {len(map_df)} field mappings")
                
                with st.spinner("Processing Extractd data..."):
                    results_df = process_Extractd_data(url_df, map_df, progress_bar3, status_text3)
                
                # Add image paths to results if enabled
                if enable_images and 'Image_src_to_path' in url_df.columns:
                    # Create mapping from Image_src to processed image path
                    image_src_to_path = {}
                    for idx, row in url_df.iterrows():
                        img_src = row.get('Image_src', '')
                        img_path = row.get('Processed_Image_Path', '')
                        if img_src and img_path:
                            image_src_to_path[img_src] = img_path
                    
                    # Map to results_df using Image_src column
                    results_df['Image_Path'] = results_df['Image_src'].map(image_src_to_path).fillna('')
                
                st.success(f"✓ Processed {len(results_df)} valid records")
                
                # Statistics
                st.markdown("### 📊 Processing Results")
                col1, col2, col3, col4 = st.columns(4)
                
                with col1:
                    st.metric("Total Records", len(results_df))
                with col2:
                    st.metric("With Prices", results_df['Listing_Price'].notna().sum())
                with col3:
                    st.metric("With SKU", results_df['SKU_Actual'].notna().sum())
                with col4:
                    if enable_images:
                        st.metric("With Images", (results_df.get('Image_Path', '') != '').sum())
                    else:
                        st.metric("Complete Data", 
                                 results_df[['SKU_Actual', 'Listing_Price', 'Desc_Text']].notna().all(axis=1).sum())
                
                # Preview
                with st.expander("👁️ Preview Processed Data", expanded=True):
                    st.dataframe(results_df.head(10), use_container_width=True)
                
                # STEP 4: Create Output Excel
                st.markdown("### 💾 Step 4: Creating Output File")
                
                output_excel_path = os.path.join(temp_dir, f"OUTPUT_{uploaded_file.name}")
                
                with pd.ExcelWriter(output_excel_path, engine='openpyxl') as writer:
                    results_df.to_excel(writer, index=False, sheet_name='Processed Data')
                    url_df.to_excel(writer, index=False, sheet_name='Extractd Data')
                    map_df.to_excel(writer, index=False, sheet_name='Mapping')
                
                st.success("✓ Excel file created")
                
                # STEP 5: Insert Images (if enabled)
                if enable_images and (results_df.get('Image_Path', '') != '').any():
                    st.markdown("### 🖼️ Step 5: Inserting Images in Excel")
                    progress_bar5 = st.progress(0)
                    status_text5 = st.empty()
                    
                    with st.spinner("Inserting images..."):
                        _, success, message = run_threaded_operation(
                            excel_image_insert_worker, output_excel_path, 
                            progress_bar5, status_text5
                        )
                    
                    if success:
                        st.success(f"✓ {message}")
                    else:
                        st.warning(f"⚠️ {message}")
                
                # STEP 6: Download
                st.markdown("### 📥 Step 6: Download Results")
                
                with open(output_excel_path, 'rb') as f:
                    excel_bytes = f.read()
                
                output_filename = f"PROCESSED_{uploaded_file.name}"
                st.download_button(
                    label="📥 Download Complete Excel File",
                    data=excel_bytes,
                    file_name=output_filename,
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    type="primary",
                    use_container_width=True
                )
                
                st.success("🎉 Complete! Download your processed file above.")
                
            except Exception as e:
                st.error(f"❌ Processing error: {str(e)}")
                st.exception(e)
            
            finally:
                # Cleanup
                try:
                    import shutil
                    shutil.rmtree(temp_dir)
                except:
                    pass
    
    else:
        st.info("👆 Please upload an Excel file to get started")
        
        with st.expander("ℹ️ How to use this tool", expanded=True):
            st.markdown("""
            **Excel File Requirements:**
            1. **Column A**: URLs to Extract
            2. **Column B**: Image_src URLs (for image download)
            3. **Sheet "Mapping"**: Field mappings (Online → Main)
            
            **Processing Steps:**
            1. ✅ Upload Excel file
            2. 🌐 Extract data from URLs
            3. 📸 Download & process images to square format (optional)
            4. 🔄 Extract structured data
            5. 🖼️ Insert images in Excel
            6. 📥 Download complete file
            
            **Output Includes:**
            - Processed data with all extracted fields
            - Original URL and Image_src columns
            - Square-formatted product images inserted in Excel
            - Raw Extractd data for reference
            """)

if __name__ == "__main__":
    main()