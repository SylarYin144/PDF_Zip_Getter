import tkinter as tk
from tkinter import filedialog, messagebox, simpledialog
# import tkinter.scrolledtext as st # GUI logging disabled
import pandas as pd
import requests
import zipfile
import os
import re 
import time
import sys 
from datetime import datetime
from bs4 import BeautifulSoup
from urllib.parse import urljoin
import json
import xml.etree.ElementTree as ET
from selenium import webdriver
from selenium.webdriver.chrome.service import Service as ChromeService
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException, NoSuchElementException

# --- Configuration Constants (Primarily for defaults now) ---
DEFAULT_SCI_HUB_MIRRORS_EXAMPLE = ["https://sci-hub.se/", "https://sci-hub.st/", "https://sci-hub.box/", "https://sci-hub.ru/", "https://sci-hub.red/"]
INTER_DOI_DELAY_SECONDS = 5 
MIRROR_SWITCH_DELAY_SECONDS = 3
STANDARD_USER_AGENT = 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/100.0.4896.127 Safari/537.36'
# --- End Configuration Constants ---

# class TextRedirector(object): # GUI logging disabled
#     def __init__(self, widget, original_stdout_ref, tag="stdout"):
#         self.widget = widget
#         self.original_stdout = original_stdout_ref
#         self.tag = tag
#     def write(self, str_):
#         self.widget.configure(state='normal')
#         self.widget.insert(tk.END, str_, (self.tag,))
#         self.widget.see(tk.END)
#         self.widget.configure(state='disabled')
#         self.original_stdout.write(str_)
#         self.original_stdout.flush()
#     def flush(self):
#         self.widget.update_idletasks()
#         self.original_stdout.flush()

# def on_log_window_close(log_window_ref, original_stdout_ref, root_tk_instance): # GUI logging disabled
#     if sys.stdout != original_stdout_ref:
#         print("Log window closed by user. Restoring original stdout.", file=original_stdout_ref)
#         sys.stdout = original_stdout_ref
#     if log_window_ref:
#         try: log_window_ref.destroy()
#         except tk.TclError: pass

def format_and_log_article_status(original_row_data, doi, title, current_article_num, total_articles, 
                                  successful_downloads_count_for_stats, # Count *after* this article's attempt
                                  mirror_attempts_details, 
                                  overall_doi_status, current_user_inter_doi_delay, is_retry=False):
    try:
        buscados_percentage = (current_article_num / total_articles) * 100 if total_articles > 0 else 0
        obtenidos_percentage = (successful_downloads_count_for_stats / total_articles) * 100 if total_articles > 0 else 0
        
        log_lines = []
        retry_prefix = "[REINTENTO] " if is_retry else ""
        prefix_applied_in_this_function = False

        # log_lines.append(f"{retry_prefix}Artículo: {current_article_num}/{total_articles} ({buscados_percentage:.2f}%)")
        # log_lines.append(f"Título: {title if title else 'N/A'}")

        # # Retrieve First Author information
        # author_val = original_row_data.get('First Author', 'N/A')
        # if author_val == 'N/A': # Fallback to "Autores" (case-insensitive)
        #     autores_keys = [k for k in original_row_data.keys() if str(k).lower() == 'autores']
        #     author_val = original_row_data.get(autores_keys[0], 'N/A') if autores_keys else 'N/A'
        # if author_val == 'N/A': # Fallback to "Authors" (case-insensitive)
        #     authors_keys_en = [k for k in original_row_data.keys() if str(k).lower() == 'authors']
        #     author_val = original_row_data.get(authors_keys_en[0], 'N/A') if authors_keys_en else 'N/A'
        # log_lines.append(f"First Author: {author_val}")

        # # Prioritize "Journal/Book" for journal title
        # journal_title_val = original_row_data.get('Journal/Book', 'N/A')

        # # If "Journal/Book" is not found (or has no value and resulted in 'N/A'),
        # # then try the old 'revista' logic as a fallback.
        # if journal_title_val == 'N/A':
        #     # Case-insensitive search for 'revista'
        #     revista_keys = [k for k in original_row_data.keys() if str(k).lower() == 'revista']
        #     journal_title_val = original_row_data.get(revista_keys[0], 'N/A') if revista_keys else 'N/A'
        # log_lines.append(f"Journal/Book: {journal_title_val}")

        # # Retrieve Publication Year information
        # pub_year_val = original_row_data.get('Publication Year', 'N/A')
        # if pub_year_val == 'N/A': # Fallback to "Fecha de publicación" (case-insensitive)
        #     fecha_pub_keys = [k for k in original_row_data.keys() if str(k).lower() == 'fecha de publicación']
        #     pub_year_val = original_row_data.get(fecha_pub_keys[0], 'N/A') if fecha_pub_keys else 'N/A'
        # if pub_year_val == 'N/A': # Fallback to "Year" (case-insensitive)
        #     year_keys = [k for k in original_row_data.keys() if str(k).lower() == 'year']
        #     pub_year_val = original_row_data.get(year_keys[0], 'N/A') if year_keys else 'N/A'
        # if pub_year_val == 'N/A': # Fallback to "Año" (case-insensitive)
        #     ano_keys = [k for k in original_row_data.keys() if str(k).lower() == 'año']
        #     pub_year_val = original_row_data.get(ano_keys[0], 'N/A') if ano_keys else 'N/A'
        # log_lines.append(f"Publication Year: {pub_year_val}")
        
        # log_lines.append(f"DOI: {doi}")

        for i, attempt in enumerate(mirror_attempts_details):
            line_prefix = ""
            if is_retry and not prefix_applied_in_this_function:
                line_prefix = retry_prefix
                prefix_applied_in_this_function = True

            mirror_url, status, reason = attempt
            try:
                domain_match = re.search(r'https?://(?:www\.)?([a-zA-Z0-9.-]+)(?:/|$)', mirror_url)
                mirror_short_name = domain_match.group(1) if domain_match else mirror_url[-20:]
            except Exception: mirror_short_name = mirror_url[-20:] 
            log_lines.append(f"{line_prefix}Intento Mirror {i+1} ({mirror_short_name}): {status}. {reason if reason else ''}".strip())

        final_doi_status_prefix = ""
        if is_retry and not prefix_applied_in_this_function:
            final_doi_status_prefix = retry_prefix
            prefix_applied_in_this_function = True
        log_lines.append(f"{final_doi_status_prefix}Artículo {overall_doi_status}")

        total_downloaded_prefix = ""
        if is_retry and not prefix_applied_in_this_function:
            total_downloaded_prefix = retry_prefix
            # prefix_applied_in_this_function = True # Not strictly needed to set here as it's the last use
        log_lines.append(f"{total_downloaded_prefix}Total Descargados (actualizado): {successful_downloads_count_for_stats}/{total_articles} ({obtenidos_percentage:.2f}%)")
        
        formatted_message = "\n".join(log_lines)
        print(f"\n{formatted_message}") 
        
        # Only print waiting message if it's not the very last article overall OR if it failed (implying more retries or end)
        # This avoids "Waiting..." after the final successful article of the main loop if no retries follow.
        is_last_article_overall = (current_article_num == total_articles and not is_retry and not failed_articles_data) # Heuristic
        
        if not (overall_doi_status == "OBTENIDO" and is_last_article_overall and not is_retry) : # Check if not last successful article
             # Check if it's not the last item in its current loop (main or retry)
            if not is_retry and current_article_num < total_articles : # Main loop, not last
                 print(f"Esperando {current_user_inter_doi_delay} segundos…\n")
            elif is_retry : # Always print for retries unless it's the very very last action. This is tricky.
                 # Let's say, if there are more retries OR if this isn't the absolute last processed item.
                 # This part of the logic for waiting message might need refinement based on exact desired flow.
                 # For now, print it if it's a retry and not the last one in the retry batch.
                 # The caller of format_and_log_article_status will handle the actual time.sleep()
                 print(f"Esperando {current_user_inter_doi_delay} segundos (post-reintento)...\n")


    except Exception as e:
        print(f"\nError al formatear log para DOI {doi}: {e}")
        print(f"Fallback Log: DOI: {doi}, Status: {overall_doi_status}\n")


def clean_filename(title):
    return re.sub(r'[\\/*?:"<>|]', '_', title)

def extract_pdf_link_from_html(article_page_url, session):
    # print(f"Extrayendo HTML de: {article_page_url}")
    try:
        response = session.get(article_page_url, timeout=30); response.raise_for_status()
        soup = BeautifulSoup(response.content, 'html.parser')
        iframe = soup.find('iframe', id='pdf')
        if iframe and iframe.get('src'):
            pdf_src = iframe['src'] #; print(f"Encontrado PDF en iframe: {pdf_src}")
            if pdf_src.startswith('//'): return 'https:' + pdf_src
            elif pdf_src.startswith('/'): return urljoin(article_page_url, pdf_src)
            return pdf_src
        embed = soup.find('embed', attrs={'type': 'application/pdf'})
        if embed and embed.get('src'):
            pdf_src = embed['src'] #; print(f"Encontrado PDF en embed: {pdf_src}")
            return urljoin(article_page_url, pdf_src)
        # print(f"No se encontró enlace PDF en iframe/embed para {article_page_url}")
        return None
    except requests.exceptions.RequestException as e: # print(f"Error al obtener HTML {article_page_url}: {e}");
        return None
    except Exception as e: # print(f"Error inesperado extrayendo PDF de {article_page_url}: {e}");
        return None

def download_from_google_scholar_old(doi, title, session):
    """
    Tries to download a PDF from Google Scholar using DOI.
    """
    # print(f"Searching Google Scholar for DOI: {doi} (Title: {title if title else 'N/A'})")
    scholar_url = f"https://scholar.google.com/scholar?q={doi}"

    try:
        headers = {'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.124 Safari/537.36'}
        response = session.get(scholar_url, headers=headers, timeout=30)
        response.raise_for_status()
        soup = BeautifulSoup(response.content, 'html.parser')

        potential_links = []
        for link_tag in soup.find_all('a', href=True):
            href = link_tag['href']
            link_text = link_tag.get_text().lower()
            if href.lower().endswith('.pdf') or '[pdf]' in link_text or 'pdf' in link_text:
                if not href.startswith('http'):
                    href = urljoin(scholar_url, href) # Ensure absolute URL

                # Basic check to avoid some common non-PDF page links that might contain "pdf"
                if 'pdf' in href.lower() and not any(x in href.lower() for x in ['view', 'download=false', 'scholar.google']):
                     potential_links.append(href)
                elif href.lower().endswith('.pdf'): # More direct .pdf links
                    potential_links.append(href)

        # De-duplicate while preserving order (important for prioritization if any)
        unique_potential_links = []
        for plink in potential_links:
            if plink not in unique_potential_links:
                unique_potential_links.append(plink)

        # print(f"Found {len(unique_potential_links)} potential PDF links on Google Scholar: {unique_potential_links}")

        for pdf_url in unique_potential_links:
            # print(f"Attempting to download PDF from: {pdf_url}")
            try:
                # Try HEAD request first (more efficient if server supports it well)
                head_response = session.head(pdf_url, headers=headers, timeout=20, allow_redirects=True)
                head_response.raise_for_status()
                content_type = head_response.headers.get('Content-Type', '').lower()

                if 'application/pdf' in content_type:
                    # print(f"HEAD request successful. Content-Type: {content_type}. Proceeding with GET.")
                    pdf_response = session.get(pdf_url, headers=headers, timeout=60, stream=True) # stream=True for large files
                    pdf_response.raise_for_status()

                    # Double check content type from GET response as well
                    get_content_type = pdf_response.headers.get('Content-Type', '').lower()
                    if 'application/pdf' in get_content_type:
                        pdf_content = pdf_response.content
                        # print(f"Successfully downloaded PDF from {pdf_url}")
                        domain_match = re.search(r'https?://(?:www\.)?([a-zA-Z0-9.-]+)(?:/|$)', pdf_url)
                        source_domain = domain_match.group(1) if domain_match else "Unknown Domain"
                        return pdf_content, f"OBTENIDO (Google Scholar via {source_domain})"
                    # else:
                        # print(f"GET request for {pdf_url} did not return PDF content-type, but: {get_content_type}")
                # else:
                    # print(f"HEAD request for {pdf_url} did not indicate PDF content-type: {content_type}")

            except requests.exceptions.HTTPError as e:
                # print(f"HTTP error when trying {pdf_url}: {e.response.status_code}")
                pass # Continue to next link
            except requests.exceptions.Timeout:
                # print(f"Timeout when trying {pdf_url}")
                pass # Continue to next link
            except requests.exceptions.RequestException as e:
                # print(f"Request error when trying {pdf_url}: {e}")
                pass # Continue to next link
            except Exception as e:
                # print(f"Unexpected error when trying {pdf_url}: {e}")
                pass # Continue to next link

        return None, f"FALLO - No PDF en Google Scholar ({scholar_url})"

    except requests.exceptions.RequestException as e:
        # print(f"Error searching Google Scholar for DOI {doi}: {e}")
        return None, f"FALLO - Error búsqueda Google Scholar ({scholar_url})"
    except Exception as e:
        # print(f"Unexpected error during Google Scholar processing for DOI {doi}: {e}")
        return None, f"FALLO - Error inesperado Google Scholar ({scholar_url})"

# This is a copy of the existing download_from_google_scholar, slightly modified for testing
# and with potential improvements.
def download_from_google_scholar(doi, title, session): # Renamed from download_from_google_scholar_fixed
    print(f"FIXED: Searching Google Scholar for DOI: {doi} (Title: {title if title else 'N/A'})")
    scholar_url = f"https://scholar.google.com/scholar?hl=en&q={doi}" # Added hl=en for consistent language

    try:
        # Using a more common and recent-looking User-Agent
        headers = {
            'User-Agent': STANDARD_USER_AGENT,
            'Accept-Language': 'en-US,en;q=0.9',
            'Accept-Encoding': 'gzip, deflate, br',
            'Accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,image/avif,image/webp,image/apng,*/*;q=0.8',
            'Connection': 'keep-alive'
        }
        # It's important to handle cookies if Google Scholar starts requiring them for search results
        # For now, session should handle basic cookies.

        response = session.get(scholar_url, headers=headers, timeout=30)
        response.raise_for_status()

        # Check if the response itself is a PDF, which can happen if Google Scholar directly serves it
        # or redirects to it.
        content_type_initial = response.headers.get('Content-Type', '').lower()
        if 'application/pdf' in content_type_initial:
            print(f"FIXED: Initial response from {scholar_url} is a PDF. Content-Type: {content_type_initial}.")
            if len(response.content) > 1000: # Basic sanity check for PDF size
                 return response.content, f"OBTENIDO (Google Scholar Direct Response - {scholar_url})"
            else:
                print(f"FIXED: Initial response was PDF, but content too small. Suspicious. Proceeding to parse.")

        soup = BeautifulSoup(response.content, 'html.parser')

        potential_links = []

        # Attempt 1: Look for common PDF link patterns
        for link_tag in soup.find_all('a', href=True):
            href = link_tag['href']
            link_text = link_tag.get_text(strip=True).lower()

            # More robust check for PDF links
            is_pdf_link = False
            if href.lower().endswith('.pdf'):
                is_pdf_link = True
            elif '[pdf]' in link_text or 'pdf' in link_text or link_tag.find(lambda tag: tag.name == 'span' and 'pdf' in tag.get_text(strip=True).lower()):
                 is_pdf_link = True

            if is_pdf_link:
                # Ensure absolute URL
                if not href.startswith('http'):
                    href = urljoin(scholar_url, href)

                # Avoid known non-PDF pages or recursive Google Scholar links more carefully
                if 'scholar.google.com' in href.lower() and not href.lower().endswith('.pdf'): # Avoid linking back to scholar unless it's a direct PDF from scholar's domain
                    continue
                if any(x in href.lower() for x in [' Morales', ' Privacy', ' Terms', ' Sign in', ' Settings', ' My Citations', ' Profiles', ' cited by', ' related articles', ' versions', ' web search', 'javascript:void(0)']): # More exclusion patterns
                    continue
                if href.endswith("#"): # Skip empty fragment links
                    continue

                potential_links.append(href)

        # Attempt 2: Look for PDF links within typical result blocks (gs_ri) and side blocks (gs_ggs)
        for result_div in soup.find_all('div', class_='gs_ri'): # Each search result item
            title_link_tag = result_div.find('h3', class_='gs_rt').find('a', href=True) if result_div.find('h3', class_='gs_rt') else None
            pdf_div = result_div.find_next_sibling('div', class_='gs_ggs') # PDF often in a side div

            if pdf_div:
                pdf_link_tag = pdf_div.find('a', href=True)
                if pdf_link_tag and pdf_link_tag['href'].lower().endswith('.pdf'):
                    href = pdf_link_tag['href']
                    if not href.startswith('http'): href = urljoin(scholar_url, href)
                    potential_links.append(href)

            if title_link_tag and title_link_tag['href'].lower().endswith('.pdf'): # If main title link is a PDF
                 href = title_link_tag['href']
                 if not href.startswith('http'): href = urljoin(scholar_url, href)
                 potential_links.append(href)


        # De-duplicate while preserving order
        unique_potential_links = []
        for plink in potential_links:
            if plink not in unique_potential_links:
                unique_potential_links.append(plink)

        print(f"FIXED: Found {len(unique_potential_links)} unique potential PDF links on Google Scholar: {unique_potential_links}")

        for pdf_url in unique_potential_links:
            print(f"FIXED: Attempting to download PDF from: {pdf_url}")
            try:
                # Use a more forgiving HEAD request or skip if it causes issues
                # Some servers might not handle HEAD well for dynamically generated PDFs
                # For now, stick to the original HEAD then GET logic but be mindful
                head_response = session.head(pdf_url, headers=headers, timeout=20, allow_redirects=True)
                head_response.raise_for_status()
                content_type = head_response.headers.get('Content-Type', '').lower()

                if 'application/pdf' in content_type:
                    print(f"FIXED: HEAD request successful for {pdf_url}. Content-Type: {content_type}. Proceeding with GET.")
                    pdf_response = session.get(pdf_url, headers=headers, timeout=60, stream=True)
                    pdf_response.raise_for_status()
                    get_content_type = pdf_response.headers.get('Content-Type', '').lower()

                    if 'application/pdf' in get_content_type:
                        pdf_content = pdf_response.content
                        if len(pdf_content) < 1000: # Check if PDF is too small (e.g. error page)
                            print(f"FIXED: PDF from {pdf_url} is very small ({len(pdf_content)} bytes). May not be valid. Skipping.")
                            continue # Try next link
                        print(f"FIXED: Successfully downloaded PDF from {pdf_url}")
                        domain_match = re.search(r'https?://(?:www\.)?([a-zA-Z0-9.-]+)(?:/|$)', pdf_url)
                        source_domain = domain_match.group(1) if domain_match else "Unknown Domain"
                        return pdf_content, f"OBTENIDO (Google Scholar via {source_domain})"
                    else:
                        print(f"FIXED: GET request for {pdf_url} did not return PDF content-type, but: {get_content_type}")
                else:
                    print(f"FIXED: HEAD request for {pdf_url} did not indicate PDF content-type: {content_type}. Trying GET anyway...")
                    # Fallback: try GET even if HEAD didn't confirm PDF, some servers are tricky
                    pdf_response = session.get(pdf_url, headers=headers, timeout=60, stream=True)
                    pdf_response.raise_for_status() # Check for HTTP errors on GET
                    get_content_type = pdf_response.headers.get('Content-Type', '').lower()
                    if 'application/pdf' in get_content_type:
                        pdf_content = pdf_response.content
                        if len(pdf_content) < 1000:
                            print(f"FIXED: PDF from {pdf_url} (after GET fallback) is very small ({len(pdf_content)} bytes). Skipping.")
                            continue
                        print(f"FIXED: Successfully downloaded PDF from {pdf_url} (after GET fallback)")
                        domain_match = re.search(r'https?://(?:www\.)?([a-zA-Z0-9.-]+)(?:/|$)', pdf_url)
                        source_domain = domain_match.group(1) if domain_match else "Unknown Domain"
                        return pdf_content, f"OBTENIDO (Google Scholar via {source_domain} - GET Fallback)"
                    else:
                        print(f"FIXED: GET fallback for {pdf_url} also did not return PDF content-type: {get_content_type}")


            except requests.exceptions.HTTPError as e:
                print(f"FIXED: HTTP error when trying {pdf_url}: {e.response.status_code if e.response else 'Unknown status'}")
            except requests.exceptions.Timeout:
                print(f"FIXED: Timeout when trying {pdf_url}")
            except requests.exceptions.RequestException as e:
                print(f"FIXED: Request error when trying {pdf_url}: {e}")
            except Exception as e:
                print(f"FIXED: Unexpected error when trying {pdf_url}: {e}")

        return None, f"FALLO - No PDF en Google Scholar ({scholar_url})"

    except requests.exceptions.RequestException as e:
        print(f"FIXED: Error searching Google Scholar for DOI {doi}: {e}")
        return None, f"FALLO - Error búsqueda Google Scholar ({scholar_url})"
    except Exception as e:
        print(f"FIXED: Unexpected error during Google Scholar processing for DOI {doi}: {e}")
        return None, f"FALLO - Error inesperado Google Scholar ({scholar_url})"

def download_with_selenium_google_scholar(driver, doi, title):
    print(f"SELENIUM GS: Searching Google Scholar for DOI: {doi} (Title: {title if title else 'N/A'})")
    scholar_url = f"https://scholar.google.com/scholar?hl=en&q={doi}"
    pdf_content = None
    status_message = f"FALLO - No PDF en Google Scholar (Selenium) ({scholar_url})"

    try:
        driver.set_page_load_timeout(90) # Further increased page load timeout
        driver.get(scholar_url)
        # Wait for the page to load and results to be present
        WebDriverWait(driver, 25).until( # Further increased wait for initial results
            EC.presence_of_element_located((By.ID, "gs_res_ccl_mid"))
        )

        # Try to find links with "[PDF]" text first - these are often direct links
        pdf_links_elements = []
        try:
            pdf_links_elements = WebDriverWait(driver, 10).until( # Increased wait
                EC.presence_of_all_elements_located((By.PARTIAL_LINK_TEXT, "[PDF]"))
            )
        except TimeoutException:
            print(f"SELENIUM GS: No direct '[PDF]' links found for {doi}. Trying other methods.")

        # If no "[PDF]" links, try to find any link containing '.pdf' in href
        if not pdf_links_elements:
            try:
                pdf_links_elements = WebDriverWait(driver, 10).until( # Increased wait
                    EC.presence_of_all_elements_located((By.XPATH, "//a[contains(@href, '.pdf')]"))
                )
            except TimeoutException:
                print(f"SELENIUM GS: No links with '.pdf' in href found for {doi}.")

        print(f"SELENIUM GS: Found {len(pdf_links_elements)} potential PDF links.")

        for link_element in pdf_links_elements:
            pdf_url = link_element.get_attribute('href')
            if pdf_url:
                print(f"SELENIUM GS: Attempting to download from URL: {pdf_url}")
                try:
                    # Use requests to download the PDF
                    # It's good practice to use a session for requests if making multiple calls,
                    # but for a single download, a direct get is fine.
                    # Ensure a User-Agent is set.
                    # Use STANDARD_USER_AGENT
                    # Create a new requests session for this download to ensure no cookie conflicts with SciHub session
                    gs_session = requests.Session()
                    gs_session.headers.update({'User-Agent': STANDARD_USER_AGENT})

                    pdf_response = gs_session.get(pdf_url, timeout=60, stream=True, allow_redirects=True)
                    pdf_response.raise_for_status() # Check for HTTP errors

                    content_type = pdf_response.headers.get('Content-Type', '').lower()
                    print(f"SELENIUM GS: URL {pdf_url} - Content-Type: {content_type}") # Added logging for content type
                    if 'application/pdf' in content_type:
                        # Check if content is substantial (more than a few KB, e.g. 1KB)
                        # Some error pages might be served as PDF
                        current_pdf_content = pdf_response.content
                        if len(current_pdf_content) > 1024: # Check if PDF is larger than 1KB
                            pdf_content = current_pdf_content
                            status_message = f"OBTENIDO (Google Scholar Selenium - {pdf_url})"
                            print(f"SELENIUM GS: Successfully downloaded PDF from {pdf_url}")
                            gs_session.close()
                            break # Exit loop once a PDF is successfully downloaded
                        else:
                            print(f"SELENIUM GS: PDF from {pdf_url} is too small ({len(current_pdf_content)} bytes). May not be valid. Trying next link.")
                    else:
                        print(f"SELENIUM GS: URL {pdf_url} is not 'application/pdf'. It is '{content_type}'. PDF might be embedded or require browser context. Skipping this link for direct requests.")
                    gs_session.close()
                except requests.exceptions.HTTPError as e_http:
                    print(f"SELENIUM GS: HTTP error downloading {pdf_url}: {e_http.response.status_code}")
                except requests.exceptions.RequestException as e_req:
                    print(f"SELENIUM GS: Request error downloading {pdf_url}: {e_req}")
                except Exception as e_gen:
                    print(f"SELENIUM GS: Unexpected error downloading {pdf_url}: {e_gen}")
            if pdf_content: # If we got content in the inner loop, break outer
                break

    except TimeoutException as e:
        # Check if it's a page load timeout vs element timeout by inspecting current_url vs scholar_url
        if driver.current_url == scholar_url or driver.current_url == "about:blank": # Heuristic: still on or trying to load the main search page
            print(f"SELENIUM GS: Page load TimeoutException for {scholar_url}: {e}")
            status_message = f"FALLO - Timeout carga página Google Scholar (Selenium) ({scholar_url})"
        else: # Timeout likely occurred waiting for an element
            print(f"SELENIUM GS: Element TimeoutException en Google Scholar (Selenium) para {doi}: {e}")
            status_message = f"FALLO - Timeout localizando elemento en Google Scholar (Selenium) ({driver.current_url})"
    except NoSuchElementException as e:
        print(f"SELENIUM GS: NoSuchElementException en Google Scholar (Selenium) para {doi}: {e}")
        status_message = f"FALLO - Elemento no encontrado en Google Scholar (Selenium) ({driver.current_url})"
    except Exception as e:
        print(f"SELENIUM GS: An unexpected error occurred with Selenium for DOI {doi} at {driver.current_url if driver else scholar_url}: {e}")
        status_message = f"FALLO - Error inesperado en Google Scholar (Selenium) ({driver.current_url if driver else scholar_url}, {str(e)[:100]})"

    return pdf_content, status_message

def download_with_selenium_pmc(driver, doi, title):
    print(f"SELENIUM PMC: Searching PubMed Central for DOI: {doi} (Title: {title if title else 'N/A'})")
    search_url = f"https://www.ncbi.nlm.nih.gov/pmc/?term={doi}"
    pdf_content = None
    status_message = f"FALLO - No PDF en PMC (Selenium) ({search_url})"
    article_url = None # To store the actual article page URL if found

    try:
        driver.set_page_load_timeout(90) # Further increased page load timeout
        driver.get(search_url)

        # Wait for search results to appear (look for a common container or result item)
        WebDriverWait(driver, 25).until( # Further increased wait
            EC.presence_of_element_located((By.CLASS_NAME, "rprt")) # Common class for search result items
        )

        # Find the first search result link that seems to be an article link
        # This might need refinement based on actual PMC search result structure
        article_link_element = None
        try:
            # Look for a link within the first result that contains the DOI or part of the title if available
            # This is a heuristic and might need adjustment.
            # Prioritize links that are clearly article links.
            possible_article_links = driver.find_elements(By.CSS_SELECTOR, "div.rprt .title a")
            if not possible_article_links: # Fallback if the above selector fails
                 possible_article_links = driver.find_elements(By.XPATH, "//div[contains(@class, 'rprt')]//a[contains(@href, 'articles/PMC')]")

            if possible_article_links:
                # For simplicity, take the first one. More complex logic could verify against title/DOI.
                article_link_element = possible_article_links[0]
                article_url = article_link_element.get_attribute('href')
                print(f"SELENIUM PMC: Found article link: {article_url}. Navigating...")
                driver.set_page_load_timeout(90) # Further increased page load timeout
                driver.get(article_url)
                print(f"SELENIUM PMC: Navigation to article page {article_url} presumably successful.")
            else:
                print(f"SELENIUM PMC: No clear article link found in search results for {doi}. Assuming current page ({driver.current_url}) might be the article page or search failed.")
                article_url = driver.current_url # Use current URL
        except Exception as e_inner_nav:
            print(f"SELENIUM PMC: Error during article link navigation for DOI {doi}: {e_inner_nav}. Proceeding with current page {driver.current_url}.")
            article_url = driver.current_url # Fallback to current URL

        # Now on the article page (or what we hope is the article page)
        # Wait for the PDF link to be present
        WebDriverWait(driver, 15).until(
            EC.presence_of_element_located((By.XPATH, "//a[contains(@href, '.pdf') or contains(translate(., 'PDF', 'pdf'), 'pdf')]"))
        )

        pdf_link_elements = []
        # Try common selectors for PDF links on PMC article pages
        selectors = [
            (By.XPATH, "//a[contains(@class, 'format-pdf') and contains(@href, '.pdf')]"), # Specific class
            (By.XPATH, "//a[contains(translate(., 'PDF', 'pdf'), 'pdf') and contains(@href, '.pdf')]"), # Contains "pdf" text and .pdf in href
            (By.PARTIAL_LINK_TEXT, "Download PDF"),
            (By.CSS_SELECTOR, "a.pdf-button[href$='.pdf']"),
            (By.XPATH, "//a[contains(@href, '.pdf') and .//span[contains(translate(., 'PDF', 'pdf'), 'pdf')]]") # Link with .pdf href and a span with "pdf"
        ]

        for by, selector_val in selectors:
            try:
                elements = WebDriverWait(driver, 5).until( # Increased wait
                    EC.presence_of_all_elements_located((by, selector_val))
                )
                if elements:
                    pdf_link_elements.extend(elements)
                    print(f"SELENIUM PMC: Found elements with selector {by} {selector_val}")
            except TimeoutException:
                print(f"SELENIUM PMC: Timeout for selector {by} {selector_val}")

        # Deduplicate elements if necessary (though order of selectors provides some priority)
        # For now, just iterate through what we found
        print(f"SELENIUM PMC: Found {len(pdf_link_elements)} potential PDF links on article page {article_url if article_url else driver.current_url}.")

        for link_element in pdf_link_elements:
            pdf_url_on_page = link_element.get_attribute('href')
            if pdf_url_on_page:
                # Ensure URL is absolute
                if not pdf_url_on_page.startswith('http'):
                    base_for_relative = driver.current_url # Base URL of the current page
                    # A more robust way to get the base for ncbi.nlm.nih.gov
                    if "ncbi.nlm.nih.gov" in base_for_relative:
                        base_for_relative = "https://www.ncbi.nlm.nih.gov"
                    pdf_url_on_page = urljoin(base_for_relative, pdf_url_on_page)

                print(f"SELENIUM PMC: Attempting to download from PDF URL: {pdf_url_on_page}")
                try:
                    # Use STANDARD_USER_AGENT
                    pmc_session = requests.Session()
                    pmc_session.headers.update({'User-Agent': STANDARD_USER_AGENT})

                    pdf_response = pmc_session.get(pdf_url_on_page, timeout=60, stream=True, allow_redirects=True)
                    pdf_response.raise_for_status()

                    content_type = pdf_response.headers.get('Content-Type', '').lower()
                    print(f"SELENIUM PMC: URL {pdf_url_on_page} - Content-Type: {content_type}") # Added logging for content type
                    if 'application/pdf' in content_type:
                        current_pdf_content = pdf_response.content
                        if len(current_pdf_content) > 1024: # Min 1KB
                            pdf_content = current_pdf_content # Assign to function-scoped variable
                            status_message = f"OBTENIDO (PMC Selenium Direct - {pdf_url_on_page})"
                            print(f"SELENIUM PMC: Successfully downloaded PDF from {pdf_url_on_page} (direct requests).")
                            # pmc_session.close() # session is closed below
                            # break # Break from this inner try, outer loop will break if pdf_content is set
                        else:
                            print(f"SELENIUM PMC: PDF from {pdf_url_on_page} (direct requests) is too small ({len(current_pdf_content)} bytes). Considered invalid.")
                    else: # Content-Type is not PDF, try Selenium navigation for embed
                        print(f"SELENIUM PMC: URL {pdf_url_on_page} returned HTML via requests. Attempting Selenium navigation to it and searching for embedded PDF.")
                        try:
                            driver.get(pdf_url_on_page)
                            WebDriverWait(driver, 20).until(lambda d: d.execute_script('return document.readyState') == 'complete')

                            embed_element = WebDriverWait(driver, 10).until(
                                EC.presence_of_element_located((By.XPATH, "//embed[@type='application/pdf']"))
                            )
                            if embed_element:
                                new_pdf_src = embed_element.get_attribute('src')
                                if new_pdf_src:
                                    print(f"SELENIUM PMC: Found embedded PDF src: {new_pdf_src}. Attempting download...")
                                    if not new_pdf_src.startswith('http'):
                                        new_pdf_src = urljoin(driver.current_url, new_pdf_src)

                                    try:
                                        embed_session = requests.Session()
                                        embed_session.headers.update({'User-Agent': STANDARD_USER_AGENT})
                                        embed_pdf_response = embed_session.get(new_pdf_src, timeout=60, stream=True, allow_redirects=True)
                                        embed_pdf_response.raise_for_status()
                                        embed_content_type = embed_pdf_response.headers.get('Content-Type', '').lower()
                                        print(f"SELENIUM PMC: Embedded URL {new_pdf_src} - Content-Type: {embed_content_type}")

                                        if 'application/pdf' in embed_content_type:
                                            current_pdf_content = embed_pdf_response.content
                                            if len(current_pdf_content) > 1024:
                                                pdf_content = current_pdf_content
                                                status_message = f"OBTENIDO (PMC Selenium Embed - {new_pdf_src})"
                                                print(f"SELENIUM PMC: Successfully downloaded PDF from embedded src: {new_pdf_src}")
                                            else:
                                                print(f"SELENIUM PMC: Embedded PDF from {new_pdf_src} is too small.")
                                        else:
                                            print(f"SELENIUM PMC: Embedded URL {new_pdf_src} is not 'application/pdf'. It is '{embed_content_type}'.")
                                        embed_session.close()
                                    except Exception as e_embed_dl:
                                        print(f"SELENIUM PMC: Error downloading embedded PDF src {new_pdf_src}: {e_embed_dl}")
                        except TimeoutException: # Timeout for driver.get(pdf_url_on_page) or finding embed
                            print(f"SELENIUM PMC: Timeout when Selenium navigated to/processed supposed PDF URL (now HTML view): {pdf_url_on_page} or no embed found.")
                        except Exception as e_nav_embed: # Other errors during navigation or embed finding
                            print(f"SELENIUM PMC: Error during Selenium navigation/embed search for {pdf_url_on_page}: {e_nav_embed}")
                    pmc_session.close() # Close session for the initial requests.get()
                except requests.exceptions.HTTPError as e_http:
                    print(f"SELENIUM PMC: HTTP error downloading {pdf_url_on_page} (initial requests): {e_http.response.status_code}")
                except requests.exceptions.RequestException as e_req:
                    print(f"SELENIUM PMC: Request error downloading {pdf_url_on_page} (initial requests): {e_req}")
                except Exception as e_gen:
                    print(f"SELENIUM PMC: Unexpected error downloading {pdf_url_on_page} (initial requests): {e_gen}")

            if pdf_content: # If PDF was obtained either directly or via embed
                break # Break from the for loop iterating through pdf_link_elements

        if not pdf_content: # If loop finishes and no PDF
             status_message = f"FALLO - No suitable PDF link found or downloaded on PMC article page ({article_url if article_url else driver.current_url})"


    except TimeoutException as e:
        current_url_for_log = driver.current_url if driver else search_url
        # Distinguish page load timeout from element location timeout
        if current_url_for_log == search_url or (article_url and current_url_for_log == article_url and not pdf_content) or current_url_for_log == "about:blank":
            print(f"SELENIUM PMC: Page load TimeoutException for {current_url_for_log} (DOI {doi}): {e}")
            status_message = f"FALLO - Timeout carga página PMC (Selenium) ({current_url_for_log})"
        else:
            print(f"SELENIUM PMC: Element TimeoutException en PMC (Selenium) for DOI {doi} at {current_url_for_log}: {e}")
            status_message = f"FALLO - Timeout localizando elemento en PMC (Selenium) ({current_url_for_log})"
    except NoSuchElementException as e:
        current_url_for_log = driver.current_url if driver else search_url
        print(f"SELENIUM PMC: NoSuchElementException en PMC (Selenium) for DOI {doi} at {current_url_for_log}: {e}")
        status_message = f"FALLO - Elemento no encontrado en PMC (Selenium) ({current_url_for_log})"
    except Exception as e:
        current_url_for_log = driver.current_url if driver else search_url
        print(f"SELENIUM PMC: An unexpected error occurred with Selenium for DOI {doi} at {current_url_for_log}: {e}")
        status_message = f"FALLO - Error inesperado en PMC (Selenium) ({current_url_for_log}, {str(e)[:100]})"

    return pdf_content, status_message

import itertools # Added for itertools.chain

def download_from_pmc(doi, title, session): # Renamed from download_from_pmc_fixed
    print(f"FIXED PMC: Attempting PubMed Central download for DOI: {doi}")
    try:
        session.headers.update({'User-Agent': STANDARD_USER_AGENT})

        id_conv_url = f"https://www.ncbi.nlm.nih.gov/pmc/utils/idconv/v1.0/?ids={doi}&format=json&tool=my_awesome_tool&email=myemail@example.com"
        try:
            response_id_conv = session.get(id_conv_url, timeout=20)
            response_id_conv.raise_for_status()
            data_id_conv = response_id_conv.json()
            # print(f"FIXED PMC: IDConv response: {data_id_conv}") # Verbose, remove for final
        except requests.exceptions.RequestException as e:
            return None, f"FALLO - Error API conversión PMCID para {doi} ({str(e)[:50]})"
        except json.JSONDecodeError:
            return None, f"FALLO - Error decodificando respuesta PMCID para {doi}"

        pmcid = None
        if data_id_conv.get("records") and len(data_id_conv["records"]) > 0:
            record = data_id_conv["records"][0]
            if record.get("pmcid"):
                 pmcid = record["pmcid"]
                 if record.get("status") == "error" and record.get("errmsg") == "invalid article id":
                      # print(f"FIXED PMC: IDConv reported 'invalid article id' for {doi} but still provided PMCID {pmcid}. Proceeding cautiously.") # Verbose
                      pass # Allow to proceed if PMCID was somehow returned despite error
            elif record.get("status") == "error":
                return None, f"FALLO - PMCID no encontrado, API devolvió '{record.get('errmsg', 'Unknown error')}' para DOI {doi}"

        if not pmcid:
            return None, f"FALLO - PMCID no encontrado para DOI {doi} (Respuesta: {data_id_conv})"

        # print(f"FIXED PMC: Found PMCID: {pmcid} for DOI: {doi}") # Verbose

        efetch_url = f"https://eutils.ncbi.nlm.nih.gov/entrez/eutils/efetch.fcgi?db=pmc&id={pmcid}&rettype=xml&tool=my_awesome_tool&email=myemail@example.com"
        try:
            # print(f"FIXED PMC: Fetching Efetch XML from: {efetch_url}") # Verbose
            response_efetch = session.get(efetch_url, timeout=45)
            response_efetch.raise_for_status()
            xml_content_bytes = response_efetch.content
            root = ET.fromstring(xml_content_bytes)

            pdf_filename_from_xml = None
            pdf_url_from_xml_constructed = None
            namespaces = {'xlink': 'http://www.w3.org/1999/xlink'}

            for tag_name_to_search in ["self-uri", "uri"]:
                combined_iterator = itertools.chain(
                    root.iterfind(f".//{tag_name_to_search}"),
                    root.iterfind(f".//{{{namespaces['xlink']}}}{tag_name_to_search}")
                )
                for element in combined_iterator:
                    # print(f"FIXED PMC: XML Efetch: Checking element <{element.tag}> with attributes {element.attrib}") # Verbose
                    content_type = element.get("content-type", "").lower()
                    href_xlink = element.get(f"{{{namespaces['xlink']}}}href")
                    href_plain = element.get("href")

                    current_href_value = None
                    if href_xlink:
                        current_href_value = href_xlink
                    elif href_plain:
                        current_href_value = href_plain

                    if "pdf" in content_type and current_href_value:
                        current_href_value = current_href_value.strip()
                        if current_href_value.lower().endswith('.pdf') or '.pdf?' in current_href_value.lower():
                            pdf_filename_from_xml = current_href_value
                            # print(f"FIXED PMC: XML Efetch: Using PDF filename/link from content-type='{content_type}': {pdf_filename_from_xml}") # Verbose
                            break
                if pdf_filename_from_xml:
                    break

            if not pdf_filename_from_xml:
                # print(f"FIXED PMC: XML Efetch: No 'content-type包含pdf' link found in self-uri/uri. Checking article-id.") # Verbose
                for element in root.iterfind(".//article-id[@pub-id-type='pmc-pdf']"):
                    if element.text:
                        href_value = element.text.strip()
                        if href_value.lower().endswith('.pdf') or '.pdf?' in href_value.lower():
                            pdf_filename_from_xml = href_value
                            # print(f"FIXED PMC: XML Efetch: Using PDF filename from article-id: {pdf_filename_from_xml}") # Verbose
                            break

            if pdf_filename_from_xml:
                if pdf_filename_from_xml.startswith('http://') or pdf_filename_from_xml.startswith('https://'):
                    pdf_url_from_xml_constructed = pdf_filename_from_xml
                elif not pdf_filename_from_xml.startswith('/'):
                    pdf_url_from_xml_constructed = f"https://www.ncbi.nlm.nih.gov/pmc/articles/{pmcid}/pdf/{pdf_filename_from_xml}"
                else:
                    pdf_url_from_xml_constructed = urljoin(f"https://www.ncbi.nlm.nih.gov/pmc/articles/{pmcid}/", pdf_filename_from_xml)
                # print(f"FIXED PMC: XML Efetch: Constructed download URL: {pdf_url_from_xml_constructed}") # Verbose

                # print(f"FIXED PMC: Attempting download from Efetch-XML-derived URL: {pdf_url_from_xml_constructed}") # Verbose
                try:
                    head_response = session.head(pdf_url_from_xml_constructed, timeout=20, allow_redirects=True)
                    head_response.raise_for_status()
                    content_type_head = head_response.headers.get('Content-Type', '').lower()
                    # print(f"FIXED PMC: Efetch HEAD from {pdf_url_from_xml_constructed} Content-Type: {content_type_head}") # Verbose

                    if 'application/pdf' in content_type_head:
                        pdf_response = session.get(pdf_url_from_xml_constructed, timeout=60, stream=True)
                        pdf_response.raise_for_status()
                        content_type_get = pdf_response.headers.get('Content-Type', '').lower()
                        # print(f"FIXED PMC: Efetch GET (after HEAD success) from {pdf_url_from_xml_constructed} Content-Type: {content_type_get}") # Verbose
                        if 'application/pdf' in content_type_get:
                            pdf_content_bytes = pdf_response.content
                            if len(pdf_content_bytes) > 1000:
                                # print(f"FIXED PMC: Successfully downloaded PDF via Efetch XML (HEAD then GET) from {pdf_url_from_xml_constructed}") # Verbose
                                return pdf_content_bytes, f"OBTENIDO (PMC Efetch XML {pmcid})"
                            # else: print(f"FIXED PMC: Efetch GET from {pdf_url_from_xml_constructed} PDF content too small.") # Verbose
                        # else: print(f"FIXED PMC: Efetch GET from {pdf_url_from_xml_constructed} did not return PDF content-type.") # Verbose
                    else:
                        # print(f"FIXED PMC: Efetch HEAD from {pdf_url_from_xml_constructed} did not indicate PDF. Trying GET anyway...") # Verbose
                        pdf_response = session.get(pdf_url_from_xml_constructed, timeout=60, stream=True)
                        pdf_response.raise_for_status()
                        content_type_get = pdf_response.headers.get('Content-Type', '').lower()
                        # print(f"FIXED PMC: Efetch GET (after HEAD failed) from {pdf_url_from_xml_constructed} Content-Type: {content_type_get}") # Verbose
                        if 'application/pdf' in content_type_get:
                            pdf_content_bytes = pdf_response.content
                            if len(pdf_content_bytes) > 1000:
                                # print(f"FIXED PMC: Successfully downloaded PDF via Efetch XML (GET fallback) from {pdf_url_from_xml_constructed}") # Verbose
                                return pdf_content_bytes, f"OBTENIDO (PMC Efetch XML - GET Fallback {pmcid})"
                            # else: print(f"FIXED PMC: Efetch GET (fallback) from {pdf_url_from_xml_constructed} PDF content too small.") # Verbose
                        # else: print(f"FIXED PMC: Efetch GET (fallback) from {pdf_url_from_xml_constructed} did not return PDF content-type.") # Verbose
                except requests.exceptions.RequestException as e_dl:
                    # print(f"FIXED PMC: Efetch XML download attempt from {pdf_url_from_xml_constructed} failed: {e_dl}") # Verbose
                    pass
            # else:
                 # print(f"FIXED PMC: XML Efetch: No suitable PDF link/filename found in XML after all checks.") # Verbose

        except requests.exceptions.RequestException: # Simplified from e_efetch for brevity in main script
            pass
        except ET.ParseError: # Simplified from e_xml
            pass

        # Step 3: Fallback Method - HTML Scraping
        # print(f"FIXED PMC: Efetch XML method failed or no PDF. Trying HTML scraping for {pmcid}.") # Verbose
        article_url = f"https://www.ncbi.nlm.nih.gov/pmc/articles/{pmcid}/"
        try:
            # print(f"FIXED PMC: Fetching HTML page: {article_url}") # Verbose
            response_html = session.get(article_url, timeout=30)
            response_html.raise_for_status()
        except requests.exceptions.RequestException as e:
            return None, f"FALLO - Error obteniendo página HTML PMC ({article_url}, {str(e)[:50]})"

        soup = BeautifulSoup(response_html.content, 'html.parser')
        potential_html_pdf_links = []
        selectors = [
            'div.format-menu a[href$=".pdf"]', 'ul.format-menu a[href$=".pdf"]',
            'div.full-text-links a[href$=".pdf"]', 'li.pdf-link a[href$=".pdf"]',
            'a.format-pdf[href$=".pdf"]', 'a.pdf-button[href$=".pdf"]', 'a.pdf-btn[href$=".pdf"]',
            'a[title*="PDF"][href$=".pdf"]', 'a[data-format="pdf"][href$=".pdf"]',
            'a[download$=".pdf"]',
            'div.buttons.article-actions a.pdf-download[href*="pdf"]'
        ]
        for selector in selectors:
            for link_tag in soup.select(selector):
                href = link_tag.get('href')
                if href: potential_html_pdf_links.append(urljoin(article_url, href.strip()))

        if not potential_html_pdf_links:
            # print("FIXED PMC: Specific selectors found no links. Trying broader search for links ending in .pdf.") # Verbose
            for link_tag in soup.find_all('a', href=lambda h: h is not None and (h.lower().endswith('.pdf') or '.pdf?' in h.lower())):
                href = link_tag.get('href')
                if pmcid.lower() in href.lower() or "articles" in href.lower() or "ftrender" in href.lower():
                     potential_html_pdf_links.append(urljoin(article_url, href.strip()))

        unique_html_pdf_links = []
        for plink in potential_html_pdf_links:
            if plink not in unique_html_pdf_links: unique_html_pdf_links.append(plink)

        # print(f"FIXED PMC: Found {len(unique_html_pdf_links)} unique potential PDF links via HTML scraping: {unique_html_pdf_links}") # Verbose

        if not unique_html_pdf_links:
            return None, f"FALLO - PDF no encontrado en HTML página PMC ({article_url})"

        for pdf_url_html in unique_html_pdf_links:
            # print(f"FIXED PMC: Attempting download from HTML-scraped URL: {pdf_url_html}") # Verbose
            try:
                pdf_response = session.get(pdf_url_html, timeout=60, stream=True)
                pdf_response.raise_for_status()
                content_type = pdf_response.headers.get('Content-Type', '').lower()
                # print(f"FIXED PMC: HTML GET from {pdf_url_html} Content-Type: {content_type}") # Verbose
                if 'application/pdf' in content_type:
                    pdf_data = pdf_response.content
                    if len(pdf_data) > 1000:
                        # print(f"FIXED PMC: Successfully downloaded PDF via HTML scraping from {pdf_url_html}") # Verbose
                        return pdf_data, f"OBTENIDO (PMC HTML {pmcid})"
                    # else: print(f"FIXED PMC: HTML GET from {pdf_url_html} PDF content too small.") # Verbose
                # else: print(f"FIXED PMC: HTML GET from {pdf_url_html} did not return PDF content-type.") # Verbose
            except requests.exceptions.RequestException: # Simplified e_html_dl
                # print(f"FIXED PMC: HTML scrape download from {pdf_url_html} failed: {e_html_dl}") # Verbose
                continue

        return None, f"FALLO - No se pudo descargar PDF desde enlaces HTML PMC ({article_url})"

    except Exception as e:
        # import traceback # Keep for debugging if needed, but remove for final script
        # print(f"FIXED PMC: Unexpected error for DOI {doi}: {e}\n{traceback.format_exc()}") # Verbose
        return None, f"FALLO - Error inesperado en PubMed Central para {doi} ({str(e)[:50]})"

def print_to_console(message, orig_stdout):
    print(message, file=orig_stdout)

def download_pdfs_from_file():
    # WebDriver will be initialized here
    driver = None  # Initialize driver to None
    try:
        options = webdriver.ChromeOptions()
        options.add_argument('--headless')
        options.add_argument('--disable-gpu')  # Recommended for headless
        options.add_argument('--no-sandbox') # Often needed in restricted environments
        options.add_argument('--disable-dev-shm-usage') # Often needed in restricted environments
        # Optional: Set a common user agent for Selenium if needed, though browser's default is usually fine
        # options.add_argument("user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.124 Safari/537.36")
        print("Inicializando WebDriver de Selenium en modo headless...")
        driver = webdriver.Chrome(service=ChromeService(ChromeDriverManager().install()), options=options)
        print("WebDriver de Selenium inicializado correctamente.")
    except Exception as e:
        print(f"Error al inicializar WebDriver de Selenium: {e}")
        print("Las descargas basadas en Selenium se omitirán.")
        driver = None # Ensure driver is None if initialization fails
    original_stdout = sys.stdout 
    # root = tk.Tk(); root.withdraw() # GUI elements removed
    # log_window = None; log_text_widget = None # GUI elements removed
    df = None # Initialize df to None
    temp_pdf_paths = [] # Initialize temp_pdf_paths as it's used in finally

    # GUI Log window and stdout redirection are disabled.
    # All print statements will go to the original stdout (console).

    try:
        # Since GUI dialogs are used, we still need a root Tk window for them,
        # but it can be managed more minimally if no log window is tied to it.
        # For now, assuming dialogs are essential and keeping root.
        root = tk.Tk()
        root.withdraw() # Keep it hidden as it's mainly for dialogs

        print("DIAG: --- Iniciando Configuración ---") # No longer using print_to_console
        input_file_path = filedialog.askopenfilename(title="Seleccionar archivo con DOIs (Excel o CSV)", filetypes=(("Archivos Excel", "*.xlsx *.xls"), ("Archivos CSV", "*.csv"), ("Todos los archivos", "*.*")))
        print(f"DIAG: After input_file_path dialog, path: {input_file_path}")
        if not input_file_path: messagebox.showinfo("Información", "No se seleccionó ningún archivo de entrada. El programa terminará."); return
        print("DIAG: Before zip_path dialog")
        zip_path = filedialog.asksaveasfilename(title="Guardar archivo ZIP como...", defaultextension=".zip", filetypes=(("Archivos ZIP", "*.zip"), ("Todos los archivos", "*.*")))
        print(f"DIAG: After zip_path dialog, path: {zip_path}")
        if not zip_path: messagebox.showinfo("Información", "No se especificó la ubicación para guardar el ZIP. El programa terminará."); return
        print("DIAG: Before excel_report_path_config dialog")
        excel_report_path_config = filedialog.asksaveasfilename(title="Guardar Reporte Excel Opcional como...", defaultextension=".xlsx", initialfile="SciHub_Descarga_Reporte.xlsx", filetypes=(("Archivos Excel", "*.xlsx"), ("Todos los archivos", "*.*")))
        print(f"DIAG: After excel_report_path_config dialog, path: {excel_report_path_config}")
        if not excel_report_path_config: excel_report_path_config = ""; print("DIAG: Ruta para reporte Excel no especificada.")
        print("DIAG: Before user_inter_doi_delay dialog")
        user_inter_doi_delay = simpledialog.askinteger("Configurar Retraso Inter-DOI", "Ingrese el tiempo de espera (segundos) entre cada DOI:", initialvalue=INTER_DOI_DELAY_SECONDS, minvalue=0 )
        print(f"DIAG: After user_inter_doi_delay dialog, value: {user_inter_doi_delay}")
        if user_inter_doi_delay is None: user_inter_doi_delay = INTER_DOI_DELAY_SECONDS; messagebox.showinfo("Información", f"Retraso Inter-DOI no modificado, se usará el predeterminado: {user_inter_doi_delay}s")
        print("DIAG: Before user_mirror_switch_delay dialog")
        user_mirror_switch_delay = simpledialog.askinteger("Configurar Retraso Cambio de Mirror", "Ingrese el tiempo de espera (segundos) al cambiar de mirror:", initialvalue=MIRROR_SWITCH_DELAY_SECONDS, minvalue=0)
        print(f"DIAG: After user_mirror_switch_delay dialog, value: {user_mirror_switch_delay}")
        if user_mirror_switch_delay is None: user_mirror_switch_delay = MIRROR_SWITCH_DELAY_SECONDS; messagebox.showinfo("Información", f"Retraso por cambio de Mirror no modificado, se usará el predeterminado: {user_mirror_switch_delay}s")
        print("DIAG: Before user_mirror_list_str dialog")
        # Use DEFAULT_SCI_HUB_MIRRORS_EXAMPLE to pre-fill the dialog
        initial_mirrors_str = ",".join(DEFAULT_SCI_HUB_MIRRORS_EXAMPLE)
        user_mirror_list_str = simpledialog.askstring(
            "Configurar Mirrors de Sci-Hub",
            "Mirrors de Sci-Hub (separados por coma):\nSe pre-cargan los valores por defecto. Puede editarlos.\nEstos serán los únicos mirrors utilizados.",
            initialvalue=initial_mirrors_str
        )
        print(f"DIAG: After user_mirror_list_str dialog, value: {user_mirror_list_str}")
        user_defined_mirrors = []
        if user_mirror_list_str is None: messagebox.showerror("Configuración Requerida", "Configuración de mirrors cancelada. Terminando."); return
        # If the user clears the dialog, user_mirror_list_str will be empty.
        # In this case, we can fall back to the default list or handle as an error.
        # For now, let's assume an empty string means the user wants no mirrors, which should be an error or handled.
        # The existing logic handles empty strip by defaulting to a single mirror, let's refine this.
        if not user_mirror_list_str.strip():
            # If user explicitly deleted all mirrors, it's likely an error or they want to cancel.
            # However, the old logic defaulted to "https://sci-hub.se/".
            # To maintain consistency while using the new default, let's use DEFAULT_SCI_HUB_MIRRORS_EXAMPLE if empty.
            messagebox.showinfo("Información de Mirrors", f"No se ingresaron mirrors o se borraron. Usando defaults: {initial_mirrors_str}")
            user_defined_mirrors = list(DEFAULT_SCI_HUB_MIRRORS_EXAMPLE) # Use a copy
        else:
            raw_mirrors = [mirror.strip() for mirror in user_mirror_list_str.split(',') if mirror.strip()]
            for mirror_url in raw_mirrors:
                if not mirror_url.startswith(("http://", "https://")): mirror_url = "https://" + mirror_url 
                if not mirror_url.endswith('/'): mirror_url += '/'
                user_defined_mirrors.append(mirror_url)
            if not user_defined_mirrors: # This case should ideally be covered by the empty string check above.
                                         # If raw_mirrors is empty due to invalid input (e.g. just commas), then error.
                messagebox.showerror("Error de Configuración", "Lista de mirrors vacía o inválida después del procesamiento. Terminando.")
                return
        # sci_hub_base_url_for_report should be set only if user_defined_mirrors is not empty.
        # The previous logic would error if user_defined_mirrors was empty here.
        # The modified logic above tries to ensure user_defined_mirrors has defaults if input is empty.
        sci_hub_base_url_for_report = user_defined_mirrors[0] if user_defined_mirrors else "N/A"


        # print_to_console("DIAG: Before root.update() after all dialogs.", original_stdout); root.update(); print_to_console("DIAG: After root.update() after all dialogs.", original_stdout) # GUI logging disabled
        # print_to_console("DIAG: All dialogs complete. Before log_window creation.", original_stdout); log_window = tk.Toplevel(root); print_to_console("DIAG: After log_window = tk.Toplevel(root)", original_stdout) # GUI logging disabled
        # log_window.title("Log de Proceso de Descarga"); log_window.geometry("800x600") # GUI logging disabled
        # print_to_console("DIAG: Before log_text_widget creation", original_stdout); log_text_widget = st.ScrolledText(log_window, wrap=tk.WORD, state='disabled'); log_text_widget.pack(padx=10, pady=10, fill=tk.BOTH, expand=True); print_to_console("DIAG: After log_text_widget creation", original_stdout) # GUI logging disabled
        # log_window.protocol("WM_DELETE_WINDOW", lambda: on_log_window_close(log_window, original_stdout, root)) # GUI logging disabled
        # print_to_console("DIAG: Before sys.stdout redirection", original_stdout); sys.stdout = TextRedirector(log_text_widget, original_stdout) ; print_to_console("DIAG: After sys.stdout redirection. Subsequent app prints go to GUI and console.", original_stdout) # GUI logging disabled
        
        print("\n--- Configuración Aplicada ---")
        print(f"Archivo de entrada: {input_file_path}")
        print(f"Archivo ZIP de salida: {zip_path}")
        if excel_report_path_config: print(f"Archivo de reporte Excel: {excel_report_path_config}")
        else: print("Reporte Excel: No se generará (ruta no especificada).")
        print(f"Retraso Inter-DOI: {user_inter_doi_delay}s")
        print(f"Retraso Cambio de Mirror: {user_mirror_switch_delay}s")
        print(f"Mirrors Sci-Hub a utilizar: {', '.join(user_defined_mirrors)}")
        print("-----------------------------------------------------\n")
        # if log_window and log_window.winfo_exists(): log_window.update_idletasks() # GUI logging disabled

        session = requests.Session(); session.headers.update({'User-Agent': STANDARD_USER_AGENT})
        all_articles_log = []; successful_articles_data = []; failed_articles_data = []; original_input_columns = []
        
        try: # File reading try block
            file_extension = os.path.splitext(input_file_path)[1].lower()
            if file_extension in ['.xlsx', '.xls']:
                try:
                    df = pd.read_excel(input_file_path)
                except Exception as e:
                    messagebox.showerror("Error de Excel", f"Error al leer Excel: {e}")
                    raise # Re-raise to be caught by the outer exception handler for file reading
            elif file_extension == '.csv': # Explicitly handle CSV
                try:
                    df = pd.read_csv(input_file_path)
                except Exception as e:
                    messagebox.showerror("Error de CSV", f"Error al leer CSV: {e}")
                    raise # Re-raise
            else:
                messagebox.showerror("Error de Archivo", f"Tipo de archivo no soportado: {file_extension}\nPor favor, use Excel o CSV.")
                # Set df to None or raise an error to ensure it's handled before use
                # Raising an error is cleaner to be caught by the existing handler.
                raise ValueError(f"Tipo de archivo no soportado: {file_extension}")

            if df is not None: # Proceed only if df was loaded
                original_input_columns = [col for col in df.columns if col not in ['DOI', 'Title']]
            else: # Should not be reached if non-supported file types raise an error
                raise ValueError("DataFrame no fue cargado correctamente.")

        except Exception as e: # Catch any exception from file reading block
            print(f"Error fatal al leer archivo de entrada: {e}")
            # if log_window and log_window.winfo_exists(): log_window.destroy() # GUI logging disabled
            messagebox.showerror("Error Crítico de Archivo", f"No se pudo leer el archivo de entrada: {e}\nEl programa terminará.")
            return # Exit if file reading fails

        # Crucial check: If df is still None here, it means file reading failed in a way not caught above,
        # or a new path was introduced. This ensures we don't proceed.
        if df is None:
            print("Error Crítico: DataFrame (df) no fue inicializado. Terminando proceso.")
            # if log_window and log_window.winfo_exists(): log_window.destroy() # GUI logging disabled
            messagebox.showerror("Error Crítico", "DataFrame no pudo ser cargado. El programa terminará.")
            return

        successful_downloads = 0; failed_downloads_summary_list = []; total_downloaded_size_bytes = 0 # temp_pdf_paths already initialized
        zip_creation_or_main_loop_error = False 

        try: # Main processing try (zip creation and DOI loops)
            with zipfile.ZipFile(zip_path, 'w', zipfile.ZIP_DEFLATED) as zf:
                total_articles = len(df) # df should be valid here
                for index, row in df.iterrows():
                    original_row_data = row.to_dict(); start_time = datetime.now()
                    doi = str(original_row_data.get('DOI', original_row_data.get('doi', ''))).strip()
                    title = str(original_row_data.get('Title', original_row_data.get('title', ''))).strip()
                    effective_title = title if title else doi
                    
                    current_article_num_for_log = index + 1
                    mirror_attempts_details_for_doi = []
                    overall_doi_status = "FALTANTE" # Default status

                    if not doi:
                        # ... (skip logic as before, but use local vars for log call)
                        # NOTE: The initial print block should be SKIPPED if DOI is empty.
                        # The existing 'continue' handles this.
                        failure_reason_for_report = "DOI vacío"; detailed_status_for_log = "Skipped_DOI_Missing"
                        overall_doi_status = "FALTANTE (DOI Vacío)"
                        # Call format_and_log_article_status here for skipped DOI
                        # successful_downloads count doesn't change for skipped
                        format_and_log_article_status(original_row_data, doi, effective_title, current_article_num_for_log, total_articles, successful_downloads, mirror_attempts_details_for_doi, overall_doi_status, user_inter_doi_delay)
                        end_time = datetime.now() # Ensure end_time for log
                        log_entry = {**original_row_data, 'Start_Time': start_time.strftime("%Y-%m-%d %H:%M:%S"), 'End_Time': end_time.strftime("%Y-%m-%d %H:%M:%S"), 'Duration_Seconds': (end_time - start_time).total_seconds(), 'Detailed_Status': detailed_status_for_log, 'Failure_Reason': failure_reason_for_report, 'Successful_Mirror': ""}
                        all_articles_log.append(log_entry); failed_articles_data.append({**original_row_data, 'Failure_Reason': failure_reason_for_report, 'Detailed_Status': detailed_status_for_log, 'original_index': index})
                        # if log_window and log_window.winfo_exists(): log_window.update_idletasks() # GUI logging disabled
                        # Separator after logging for a skipped DOI
                        print_to_console("===============================================================================================", original_stdout)
                        time.sleep(user_inter_doi_delay) # Still sleep for skipped
                        continue
                    
                    pdf_filename_in_zip = clean_filename(effective_title)[:150] + ".pdf"

                    # --- Initial Article Summary Print Block (Main Loop) ---
                    current_article_num = index + 1 # Already have current_article_num_for_log, can reuse
                    buscados_percentage = (current_article_num / total_articles) * 100 if total_articles > 0 else 0

                    author_val_initial = original_row_data.get('First Author', 'N/A')
                    if author_val_initial == 'N/A':
                        autores_keys_initial = [k for k in original_row_data.keys() if str(k).lower() == 'autores']
                        author_val_initial = original_row_data.get(autores_keys_initial[0], 'N/A') if autores_keys_initial else 'N/A'
                    if author_val_initial == 'N/A':
                        authors_keys_en_initial = [k for k in original_row_data.keys() if str(k).lower() == 'authors']
                        author_val_initial = original_row_data.get(authors_keys_en_initial[0], 'N/A') if authors_keys_en_initial else 'N/A'

                    journal_title_val_initial = original_row_data.get('Journal/Book', 'N/A')
                    if journal_title_val_initial == 'N/A':
                        revista_keys_initial = [k for k in original_row_data.keys() if str(k).lower() == 'revista']
                        journal_title_val_initial = original_row_data.get(revista_keys_initial[0], 'N/A') if revista_keys_initial else 'N/A'

                    pub_year_val_initial = original_row_data.get('Publication Year', 'N/A')
                    if pub_year_val_initial == 'N/A':
                        fecha_pub_keys_initial = [k for k in original_row_data.keys() if str(k).lower() == 'fecha de publicación']
                        pub_year_val_initial = original_row_data.get(fecha_pub_keys_initial[0], 'N/A') if fecha_pub_keys_initial else 'N/A'
                    if pub_year_val_initial == 'N/A':
                        year_keys_initial = [k for k in original_row_data.keys() if str(k).lower() == 'year']
                        pub_year_val_initial = original_row_data.get(year_keys_initial[0], 'N/A') if year_keys_initial else 'N/A'
                    if pub_year_val_initial == 'N/A':
                        ano_keys_initial = [k for k in original_row_data.keys() if str(k).lower() == 'año']
                        pub_year_val_initial = original_row_data.get(ano_keys_initial[0], 'N/A') if ano_keys_initial else 'N/A'

                    print(f"Artículo: {current_article_num}/{total_articles} ({buscados_percentage:.2f}%)")
                    print(f"Título: {effective_title if effective_title else 'N/A'}")
                    print(f"First Author: {author_val_initial}")
                    print(f"Journal/Book: {journal_title_val_initial}")
                    print(f"Publication Year: {pub_year_val_initial}")
                    print(f"DOI: {doi}")
                    print("-" * 30)
                    # --- End Initial Article Summary Print Block ---
                    
                    mirrors_to_try_for_this_doi = list(user_defined_mirrors)
                    pdf_content = None; download_successful_this_doi = False; successful_mirror_for_this_doi = ""
                    temp_detailed_status_for_log = ""; temp_failure_reason_for_log = ""

                    for mirror_idx, current_mirror_base_url in enumerate(mirrors_to_try_for_this_doi):
                        # Print for mirror attempt will be part of format_and_log_article_status details
                        full_sci_hub_url_for_html_page = f"{current_mirror_base_url}{doi}"
                        mirror_status_str = "FALLO"; mirror_reason_str = ""
                        
                        actual_pdf_download_url = extract_pdf_link_from_html(full_sci_hub_url_for_html_page, session)
                        if actual_pdf_download_url:
                            try:
                                response = session.get(actual_pdf_download_url, timeout=60); response.raise_for_status()
                                content_type = response.headers.get('Content-Type', '').lower()
                                if 'application/pdf' in content_type: 
                                    pdf_content = response.content; mirror_status_str = "OBTENIDO (Extracción Iframe/Embed)"; temp_detailed_status_for_log = f"Success_iframe_or_embed_extraction_from_{current_mirror_base_url}"
                                else: 
                                    mirror_reason_str = f"Content-Type no PDF ({content_type})"; temp_detailed_status_for_log = f"Failure_iframe_or_embed_extraction_not_pdf_from_{current_mirror_base_url}"
                            except requests.exceptions.HTTPError as e: mirror_reason_str = f"HTTPError {e.response.status_code}"; temp_detailed_status_for_log = f"Failure_iframe_or_embed_extraction_HTTP{e.response.status_code}_from_{current_mirror_base_url}"
                            except requests.exceptions.RequestException as e: mirror_reason_str = "Error de conexión/RequestException en extracción"; temp_detailed_status_for_log = f"Failure_iframe_or_embed_extraction_RequestException_from_{current_mirror_base_url}"
                            except Exception as e: mirror_reason_str = "Error inesperado en extracción"; temp_detailed_status_for_log = f"Failure_iframe_or_embed_extraction_Unexpected_from_{current_mirror_base_url}"
                        else: mirror_reason_str = "No se encontró enlace PDF en HTML"; temp_detailed_status_for_log = f"Failure_No_PDF_Link_Found_In_HTML_from_{current_mirror_base_url}"
                        
                        if not pdf_content: # Fallback if extraction failed
                            temp_failure_reason_for_log = mirror_reason_str # Keep reason from extraction attempt
                            try:
                                response = session.get(full_sci_hub_url_for_html_page, timeout=30); response.raise_for_status()
                                content_type = response.headers.get('Content-Type', '').lower()
                                if 'application/pdf' in content_type: 
                                    pdf_content = response.content; mirror_status_str = "OBTENIDO (Acceso Directo Fallback)"; mirror_reason_str = ""; temp_detailed_status_for_log = f"Success_direct_DOI_access_fallback_from_{current_mirror_base_url}"
                                else: 
                                    mirror_reason_str = f"Content-Type no PDF ({content_type}) en acceso directo"; temp_detailed_status_for_log = f"Failure_direct_DOI_access_not_pdf_from_{current_mirror_base_url}"
                            except requests.exceptions.HTTPError as e: mirror_reason_str = f"HTTPError {e.response.status_code} en acceso directo"; temp_detailed_status_for_log = f"Failure_direct_DOI_access_HTTP{e.response.status_code}_from_{current_mirror_base_url}"
                            except requests.exceptions.RequestException as e: mirror_reason_str = "Error de conexión/RequestException en fallback"; temp_detailed_status_for_log = f"Failure_direct_DOI_access_RequestException_from_{current_mirror_base_url}"
                            except Exception as e: mirror_reason_str = "Error inesperado en fallback"; temp_detailed_status_for_log = f"Failure_direct_DOI_access_Unexpected_from_{current_mirror_base_url}"
                        
                        # Logic for appending to mirror_attempts_details_for_doi and setting temp_failure_reason_for_log
                        specific_reason_for_temp_log = mirror_reason_str # Capture specific reason from this mirror

                        log_display_reason_for_sci_hub_attempt = mirror_reason_str
                        if mirror_status_str == "FALLO":
                            log_display_reason_for_sci_hub_attempt = full_sci_hub_url_for_html_page
                            temp_failure_reason_for_log = specific_reason_for_temp_log # Update overall DOI failure with specific reason from this mirror
                        else: # OBTENIDO from this mirror
                            temp_failure_reason_for_log = "" # Clear overall DOI failure reason

                        mirror_attempts_details_for_doi.append((current_mirror_base_url, mirror_status_str, log_display_reason_for_sci_hub_attempt))

                        if pdf_content: # Successfully got PDF from current_mirror_base_url
                            download_successful_this_doi = True
                            successful_mirror_for_this_doi = current_mirror_base_url
                            # overall_doi_status will be set to "OBTENIDO" or the more specific success message from mirror_status_str
                            if mirror_status_str.startswith("OBTENIDO"): overall_doi_status = mirror_status_str
                            else: overall_doi_status = "OBTENIDO"
                            temp_failure_reason_for_log = "" # Clear overall failure reason for the DOI
                            # temp_detailed_status_for_log is already set by the successful extraction/fallback
                            break # Break from Sci-Hub mirror loop
                        else:
                            # PDF not found with this mirror, temp_failure_reason_for_log has the specific reason
                            if mirror_idx < len(mirrors_to_try_for_this_doi) - 1:
                                time.sleep(user_mirror_switch_delay)

                    # After Sci-Hub loop, if still no pdf_content, try Google Scholar
                    if not pdf_content:
                        print(f"INFO: DOI {doi} - Attempting Google Scholar (direct request method)...")
                        # print(f"Sci-Hub attempts failed for DOI {doi}. Trying Google Scholar.") # Debug print commented out
                        gs_pdf_content, gs_status_msg = download_from_google_scholar(doi, effective_title, session)
                        if gs_pdf_content:
                            pdf_content = gs_pdf_content
                            download_successful_this_doi = True
                            successful_mirror_for_this_doi = "Google Scholar"
                            overall_doi_status = gs_status_msg
                            mirror_attempts_details_for_doi.append(("Google Scholar", "OBTENIDO", gs_status_msg))
                            temp_detailed_status_for_log = f"Success_GoogleScholar_{gs_status_msg}"
                            temp_failure_reason_for_log = "" # Clear overall DOI failure as GS succeeded
                        else:
                            mirror_attempts_details_for_doi.append(("Google Scholar", "FALLO", gs_status_msg))
                            # temp_failure_reason_for_log is already set if all Sci-Hub mirrors failed.
                            # If Sci-Hub mirrors didn't run (e.g., empty list), then gs_status_msg becomes the failure reason.
                            # If Sci-Hub mirrors ran and failed, gs_status_msg will overwrite the last Sci-Hub specific error.
                            temp_failure_reason_for_log = gs_status_msg
                            temp_detailed_status_for_log = f"Failure_GoogleScholar_{gs_status_msg}"

                    if not pdf_content and driver: # Check if driver was initialized
                        print(f"INFO: DOI {doi} - Google Scholar (direct) failed. Attempting Google Scholar (Selenium method)...")
                        gs_selenium_pdf_content, gs_selenium_status_msg = download_with_selenium_google_scholar(driver, doi, effective_title)
                        if gs_selenium_pdf_content:
                            pdf_content = gs_selenium_pdf_content
                            download_successful_this_doi = True
                            successful_mirror_for_this_doi = "Google Scholar (Selenium)"
                            overall_doi_status = gs_selenium_status_msg
                            mirror_attempts_details_for_doi.append(("Google Scholar (Selenium)", "OBTENIDO", gs_selenium_status_msg))
                            temp_detailed_status_for_log = f"Success_GoogleScholar_Selenium_{gs_selenium_status_msg}"
                            temp_failure_reason_for_log = ""
                        else:
                            mirror_attempts_details_for_doi.append(("Google Scholar (Selenium)", "FALLO", gs_selenium_status_msg))
                            temp_failure_reason_for_log = gs_selenium_status_msg
                            temp_detailed_status_for_log = f"Failure_GoogleScholar_Selenium_{gs_selenium_status_msg}"

                    # After Google Scholar attempt, if still no pdf_content, try PubMed Central
                    if not pdf_content:
                        print(f"INFO: DOI {doi} - Attempting PubMed Central (direct API/scrape method)...")
                        # Optional: print(f"Sci-Hub and Google Scholar failed for {doi}. Trying PubMed Central.")
                        pmc_pdf_content, pmc_status_msg = download_from_pmc(doi, effective_title, session)
                        if pmc_pdf_content:
                            pdf_content = pmc_pdf_content
                            download_successful_this_doi = True
                            successful_mirror_for_this_doi = pmc_status_msg # e.g., "OBTENIDO (PubMed Central PMCID_HERE)"
                            overall_doi_status = pmc_status_msg
                            mirror_attempts_details_for_doi.append(("PubMed Central", "OBTENIDO", pmc_status_msg))
                            temp_detailed_status_for_log = f"Success_PubMedCentral_{pmc_status_msg}"
                            temp_failure_reason_for_log = "" # Clear failure reason as PMC succeeded
                        else:
                            mirror_attempts_details_for_doi.append(("PubMed Central", "FALLO", pmc_status_msg))
                            temp_failure_reason_for_log = pmc_status_msg # Update with PMC failure reason
                            temp_detailed_status_for_log = f"Failure_PubMedCentral_{pmc_status_msg}"

                    if not pdf_content and driver: # Check if driver was initialized
                        print(f"INFO: DOI {doi} - PubMed Central (direct) failed. Attempting PubMed Central (Selenium method)...")
                        pmc_selenium_pdf_content, pmc_selenium_status_msg = download_with_selenium_pmc(driver, doi, effective_title)
                        if pmc_selenium_pdf_content:
                            pdf_content = pmc_selenium_pdf_content
                            download_successful_this_doi = True
                            successful_mirror_for_this_doi = pmc_selenium_status_msg
                            overall_doi_status = pmc_selenium_status_msg
                            mirror_attempts_details_for_doi.append(("PubMed Central (Selenium)", "OBTENIDO", pmc_selenium_status_msg))
                            temp_detailed_status_for_log = f"Success_PubMedCentral_Selenium_{pmc_selenium_status_msg}"
                            temp_failure_reason_for_log = ""
                        else:
                            mirror_attempts_details_for_doi.append(("PubMed Central (Selenium)", "FALLO", pmc_selenium_status_msg))
                            temp_failure_reason_for_log = pmc_selenium_status_msg
                            temp_detailed_status_for_log = f"Failure_PubMedCentral_Selenium_{pmc_selenium_status_msg}"

                    end_time = datetime.now()
                    if download_successful_this_doi and pdf_content:
                        successful_downloads += 1 # Increment *before* calling log for current stats
                        # overall_doi_status is already set if successful (either by SciHub or GS)
                        if not overall_doi_status.startswith("OBTENIDO"): # Ensure it's marked OBTENIDO if somehow missed
                            overall_doi_status = "OBTENIDO"
                        # ... (save PDF to zip as before) ...
                        data_for_successful_sheet = original_row_data.copy(); data_for_successful_sheet['Successful_Mirror'] = successful_mirror_for_this_doi; successful_articles_data.append(data_for_successful_sheet)
                        temp_dir = "temp_scihub_pdfs"; 
                        if not os.path.exists(temp_dir): os.makedirs(temp_dir)
                        temp_pdf_path = os.path.join(temp_dir, f"temp_{os.getpid()}_{index}_{pdf_filename_in_zip}")
                        with open(temp_pdf_path, 'wb') as f: f.write(pdf_content)
                        try: total_downloaded_size_bytes += os.path.getsize(temp_pdf_path)
                        except OSError as e: print(f"Advertencia: tamaño temp {temp_pdf_path}: {e}") # This print will go to log
                        zf.write(temp_pdf_path, arcname=pdf_filename_in_zip); temp_pdf_paths.append(temp_pdf_path)
                    else:
                        overall_doi_status = "FALTANTE"
                        # temp_failure_reason_for_log and temp_detailed_status_for_log will have the last failure
                        failed_articles_data.append({**original_row_data, 'Failure_Reason': temp_failure_reason_for_log, 'Detailed_Status': temp_detailed_status_for_log, 'original_index': index}) # Store original index

                    format_and_log_article_status(original_row_data, doi, effective_title, current_article_num_for_log, total_articles, successful_downloads, mirror_attempts_details_for_doi, overall_doi_status, user_inter_doi_delay)
                    
                    log_entry_failure_reason = temp_failure_reason_for_log if not download_successful_this_doi else ""
                    log_entry_detailed_status = temp_detailed_status_for_log if not download_successful_this_doi else f"Success_{successful_mirror_for_this_doi}" # Simplified success status for log

                    all_articles_log.append({**original_row_data, 'Start_Time': start_time.strftime("%Y-%m-%d %H:%M:%S"), 'End_Time': end_time.strftime("%Y-%m-%d %H:%M:%S"), 'Duration_Seconds': (end_time - start_time).total_seconds(), 'Detailed_Status': log_entry_detailed_status, 'Failure_Reason': log_entry_failure_reason, 'Successful_Mirror': successful_mirror_for_this_doi })
                    
                    # if log_window and log_window.winfo_exists(): log_window.update_idletasks() # GUI logging disabled
                    # Separator after logging for the current article in the main loop
                    print_to_console("===============================================================================================", original_stdout)
                    if current_article_num_for_log < total_articles : time.sleep(user_inter_doi_delay) # Sleep if not the last article
                    # Ensure overall_doi_status is updated for the log entry if it was a GS success
                    if download_successful_this_doi and "Google Scholar" in successful_mirror_for_this_doi:
                         log_entry_detailed_status = f"Success_{successful_mirror_for_this_doi}" # Update for GS success

        # ... (except FileNotFoundError, Exception for zip as before) ...
        except FileNotFoundError: messagebox.showerror("Error", f"No se pudo crear ZIP (Directorio no encontrado): {zip_path}"); print("Error crítico: FileNotFoundError al crear ZIP."); zip_creation_or_main_loop_error = True
        except Exception as e: messagebox.showerror("Error", f"Error inesperado en ZIP o descargas: {e}"); print(f"Error crítico: Excepción en ZIP o descargas: {e}."); zip_creation_or_main_loop_error = True


        if not zip_creation_or_main_loop_error:
            if failed_articles_data: 
                print("\n--- Iniciando Fase de Reintento para Artículos Fallidos ---")
                # if log_window and log_window.winfo_exists(): log_window.update_idletasks() # GUI logging disabled
            
            articles_successfully_retried_ids = [] 
            temp_failed_articles_data_for_iteration = list(failed_articles_data) 
            mirrors_for_retry = list(user_defined_mirrors)

            for retry_idx, failed_article_entry in enumerate(temp_failed_articles_data_for_iteration):
                original_index_for_retry = failed_article_entry.get('original_index', -1) # Get original index
                current_article_num_for_log_retry = original_index_for_retry + 1 if original_index_for_retry != -1 else retry_idx + 1 # Use original index for overall progress

                doi_to_retry = str(failed_article_entry.get('DOI', failed_article_entry.get('doi', ''))).strip()
                effective_title_for_retry = str(failed_article_entry.get('Title', failed_article_entry.get('title', doi_to_retry))).strip() or doi_to_retry
                pdf_filename_in_zip_retry = clean_filename(effective_title_for_retry)[:150] + ".pdf"

                # --- Initial Article Summary Print Block (Retry Loop) ---
                # current_article_num_for_log_retry and total_articles are already available
                buscados_percentage_retry = (current_article_num_for_log_retry / total_articles) * 100 if total_articles > 0 else 0

                # Extract author, journal, year from failed_article_entry (which is original_row_data for this item)
                author_val_retry = failed_article_entry.get('First Author', 'N/A')
                if author_val_retry == 'N/A':
                    autores_keys_retry = [k for k in failed_article_entry.keys() if str(k).lower() == 'autores']
                    author_val_retry = failed_article_entry.get(autores_keys_retry[0], 'N/A') if autores_keys_retry else 'N/A'
                if author_val_retry == 'N/A':
                    authors_keys_en_retry = [k for k in failed_article_entry.keys() if str(k).lower() == 'authors']
                    author_val_retry = failed_article_entry.get(authors_keys_en_retry[0], 'N/A') if authors_keys_en_retry else 'N/A'

                journal_title_val_retry = failed_article_entry.get('Journal/Book', 'N/A')
                if journal_title_val_retry == 'N/A':
                    revista_keys_retry = [k for k in failed_article_entry.keys() if str(k).lower() == 'revista']
                    journal_title_val_retry = failed_article_entry.get(revista_keys_retry[0], 'N/A') if revista_keys_retry else 'N/A'

                pub_year_val_retry = failed_article_entry.get('Publication Year', 'N/A')
                if pub_year_val_retry == 'N/A':
                    fecha_pub_keys_retry = [k for k in failed_article_entry.keys() if str(k).lower() == 'fecha de publicación']
                    pub_year_val_retry = failed_article_entry.get(fecha_pub_keys_retry[0], 'N/A') if fecha_pub_keys_retry else 'N/A'
                if pub_year_val_retry == 'N/A':
                    year_keys_retry = [k for k in failed_article_entry.keys() if str(k).lower() == 'year']
                    pub_year_val_retry = failed_article_entry.get(year_keys_retry[0], 'N/A') if year_keys_retry else 'N/A'
                if pub_year_val_retry == 'N/A':
                    ano_keys_retry = [k for k in failed_article_entry.keys() if str(k).lower() == 'año']
                    pub_year_val_retry = failed_article_entry.get(ano_keys_retry[0], 'N/A') if ano_keys_retry else 'N/A'

                print(f"[REINTENTO] Artículo: {current_article_num_for_log_retry}/{total_articles} ({buscados_percentage_retry:.2f}%)")
                print(f"[REINTENTO] Título: {effective_title_for_retry if effective_title_for_retry else 'N/A'}")
                print(f"[REINTENTO] First Author: {author_val_retry}")
                print(f"[REINTENTO] Journal/Book: {journal_title_val_retry}")
                print(f"[REINTENTO] Publication Year: {pub_year_val_retry}")
                print(f"[REINTENTO] DOI: {doi_to_retry}")
                print("-" * 30)
                # --- End Initial Article Summary Print Block (Retry Loop) ---
                
                mirror_attempts_details_for_retry = []
                overall_retry_status = "FALTANTE"
                pdf_content_retry = None; retry_successful_this_doi = False; successful_mirror_for_retry = ""
                temp_detailed_status_for_retry_log = ""; temp_failure_reason_for_retry_log = ""
                
                retry_start_time_actual_attempt = datetime.now() # For duration calculation of this specific retry in log

                for mirror_idx_retry, current_mirror_base_url_retry in enumerate(mirrors_for_retry):
                    # ... (retry mirror attempt logic, populate mirror_attempts_details_for_retry) ...
                    # ... (similar to main loop's mirror logic, use _retry suffixed variables)
                    full_sci_hub_url_for_html_page_retry = f"{current_mirror_base_url_retry}{doi_to_retry}"
                    mirror_status_str_retry = "FALLO"; mirror_reason_str_retry = ""
                    actual_pdf_download_url_retry = extract_pdf_link_from_html(full_sci_hub_url_for_html_page_retry, session)
                    if actual_pdf_download_url_retry:
                        try: # ... (extraction)
                            response = session.get(actual_pdf_download_url_retry, timeout=60); response.raise_for_status()
                            content_type = response.headers.get('Content-Type', '').lower()
                            if 'application/pdf' in content_type: pdf_content_retry = response.content; mirror_status_str_retry = "OBTENIDO (REINTENTO Extracción)"; temp_detailed_status_for_retry_log = f"Success_RETRY_iframe_embed_from_{current_mirror_base_url_retry}"
                            else: mirror_reason_str_retry = f"RETRY: Content-Type not PDF ({content_type})"; temp_detailed_status_for_retry_log = f"Failure_RETRY_iframe_embed_not_pdf_from_{current_mirror_base_url_retry}"
                        except requests.exceptions.HTTPError as e: mirror_reason_str_retry = f"RETRY: HTTPError {e.response.status_code}"; temp_detailed_status_for_retry_log = f"Failure_RETRY_iframe_embed_HTTPError_from_{current_mirror_base_url_retry}"
                        except requests.exceptions.RequestException as e: mirror_reason_str_retry = "RETRY: Error de conexión/RequestException en extracción"; temp_detailed_status_for_retry_log = f"Failure_RETRY_iframe_embed_RequestException_from_{current_mirror_base_url_retry}"
                        except Exception as e: mirror_reason_str_retry = "RETRY: Error inesperado en extracción"; temp_detailed_status_for_retry_log = f"Failure_RETRY_iframe_embed_Unexpected_from_{current_mirror_base_url_retry}"
                    else: mirror_reason_str_retry = "RETRY: No se encontró enlace PDF en HTML"; temp_detailed_status_for_retry_log = f"Failure_RETRY_No_PDF_Link_Found_In_HTML_from_{current_mirror_base_url_retry}"

                    if not pdf_content_retry: # Fallback
                        temp_failure_reason_for_retry_log = mirror_reason_str_retry # Preserve reason from extraction if it failed there
                        try:
                            response = session.get(full_sci_hub_url_for_html_page_retry, timeout=30); response.raise_for_status()
                            content_type = response.headers.get('Content-Type', '').lower()
                            if 'application/pdf' in content_type:
                                pdf_content_retry = response.content; mirror_status_str_retry = "OBTENIDO (REINTENTO Fallback Directo)"; mirror_reason_str_retry = ""; temp_detailed_status_for_retry_log = f"Success_RETRY_direct_DOI_from_{current_mirror_base_url_retry}"
                            else:
                                mirror_reason_str_retry = f"RETRY: Content-Type not PDF ({content_type}) en fallback"; temp_detailed_status_for_retry_log = f"Failure_RETRY_direct_DOI_not_pdf_from_{current_mirror_base_url_retry}"
                        except requests.exceptions.HTTPError as e: mirror_reason_str_retry = f"RETRY: HTTPError {e.response.status_code} en fallback"; temp_detailed_status_for_retry_log = f"Failure_RETRY_direct_DOI_HTTPError_from_{current_mirror_base_url_retry}"
                        except requests.exceptions.RequestException as e: mirror_reason_str_retry = "RETRY: Error de conexión/RequestException en fallback"; temp_detailed_status_for_retry_log = f"Failure_RETRY_direct_DOI_RequestException_from_{current_mirror_base_url_retry}"
                        except Exception as e: mirror_reason_str_retry = "RETRY: Error inesperado en fallback"; temp_detailed_status_for_retry_log = f"Failure_RETRY_direct_DOI_Unexpected_from_{current_mirror_base_url_retry}"

                    # Logic for appending to mirror_attempts_details_for_retry and setting temp_failure_reason_for_retry_log
                    specific_reason_for_temp_retry_log = mirror_reason_str_retry # Capture specific reason

                    log_display_reason_for_sci_hub_retry_attempt = mirror_reason_str_retry
                    if mirror_status_str_retry == "FALLO":
                        log_display_reason_for_sci_hub_retry_attempt = full_sci_hub_url_for_html_page_retry
                        temp_failure_reason_for_retry_log = specific_reason_for_temp_retry_log
                    else: # OBTENIDO from this mirror
                        temp_failure_reason_for_retry_log = ""

                    mirror_attempts_details_for_retry.append((current_mirror_base_url_retry, mirror_status_str_retry, log_display_reason_for_sci_hub_retry_attempt))

                    if pdf_content_retry: # Successfully got PDF from current_mirror_base_url_retry
                        retry_successful_this_doi = True
                        successful_mirror_for_retry = current_mirror_base_url_retry
                        if mirror_status_str_retry.startswith("OBTENIDO"): overall_retry_status = mirror_status_str_retry
                        else: overall_retry_status = "OBTENIDO"
                        temp_failure_reason_for_retry_log = ""
                        # temp_detailed_status_for_retry_log is already set
                        break # Break from Sci-Hub retry mirror loop
                    else:
                        # PDF not found with this mirror, temp_failure_reason_for_retry_log has the specific reason
                        if mirror_idx_retry < len(mirrors_for_retry) - 1:
                            time.sleep(user_mirror_switch_delay)

                # After Sci-Hub retry loop, if still no pdf_content_retry, try Google Scholar
                if not pdf_content_retry:
                    print(f"INFO: DOI {doi_to_retry} [RETRY] - Attempting Google Scholar (direct request method)...")
                    # print(f"Sci-Hub retry attempts failed for DOI {doi_to_retry}. Trying Google Scholar.") # Debug print commented out
                    gs_pdf_content_retry, gs_status_msg_retry = download_from_google_scholar(doi_to_retry, effective_title_for_retry, session)
                    if gs_pdf_content_retry:
                        pdf_content_retry = gs_pdf_content_retry
                        retry_successful_this_doi = True
                        successful_mirror_for_retry = "Google Scholar"
                        overall_retry_status = gs_status_msg_retry
                        mirror_attempts_details_for_retry.append(("Google Scholar (Retry)", "OBTENIDO", gs_status_msg_retry))
                        temp_detailed_status_for_retry_log = f"Success_RETRY_GoogleScholar_{gs_status_msg_retry}"
                        temp_failure_reason_for_retry_log = ""
                    else:
                        mirror_attempts_details_for_retry.append(("Google Scholar (Retry)", "FALLO", gs_status_msg_retry))
                        temp_failure_reason_for_retry_log = gs_status_msg_retry
                        temp_detailed_status_for_retry_log = f"Failure_RETRY_GoogleScholar_{gs_status_msg_retry}"

                if not pdf_content_retry and driver: # Check if driver was initialized
                    print(f"INFO: DOI {doi_to_retry} [RETRY] - Google Scholar (direct) failed. Attempting Google Scholar (Selenium method)...")
                    gs_selenium_pdf_content_retry, gs_selenium_status_msg_retry = download_with_selenium_google_scholar(driver, doi_to_retry, effective_title_for_retry)
                    if gs_selenium_pdf_content_retry:
                        pdf_content_retry = gs_selenium_pdf_content_retry
                        retry_successful_this_doi = True
                        successful_mirror_for_retry = "Google Scholar (Selenium)"
                        overall_retry_status = gs_selenium_status_msg_retry
                        mirror_attempts_details_for_retry.append(("Google Scholar (Selenium Retry)", "OBTENIDO", gs_selenium_status_msg_retry))
                        temp_detailed_status_for_retry_log = f"Success_RETRY_GoogleScholar_Selenium_{gs_selenium_status_msg_retry}"
                        temp_failure_reason_for_retry_log = ""
                    else:
                        mirror_attempts_details_for_retry.append(("Google Scholar (Selenium Retry)", "FALLO", gs_selenium_status_msg_retry))
                        temp_failure_reason_for_retry_log = gs_selenium_status_msg_retry
                        temp_detailed_status_for_retry_log = f"Failure_RETRY_GoogleScholar_Selenium_{gs_selenium_status_msg_retry}"

                # If Sci-Hub retry and Google Scholar retry failed, try PubMed Central for retry
                if not pdf_content_retry:
                    print(f"INFO: DOI {doi_to_retry} [RETRY] - Attempting PubMed Central (direct API/scrape method)...")
                    # print(f"Sci-Hub & Google Scholar retry attempts failed for DOI {doi_to_retry}. Trying PubMed Central.") # Debug print
                    pmc_pdf_content_retry, pmc_status_msg_retry = download_from_pmc(doi_to_retry, effective_title_for_retry, session)
                    if pmc_pdf_content_retry:
                        pdf_content_retry = pmc_pdf_content_retry
                        retry_successful_this_doi = True
                        successful_mirror_for_retry = "PubMed Central"
                        overall_retry_status = pmc_status_msg_retry
                        mirror_attempts_details_for_retry.append(("PubMed Central (Retry)", "OBTENIDO", pmc_status_msg_retry))
                        temp_detailed_status_for_retry_log = f"Success_RETRY_PMC_{pmc_status_msg_retry}"
                        temp_failure_reason_for_retry_log = "" # Clear overall failure
                    else:
                        mirror_attempts_details_for_retry.append(("PubMed Central (Retry)", "FALLO", pmc_status_msg_retry))
                        temp_failure_reason_for_retry_log = pmc_status_msg_retry
                        temp_detailed_status_for_retry_log = f"Failure_RETRY_PubMedCentral_{pmc_status_msg_retry}"

                if not pdf_content_retry and driver: # Check if driver was initialized
                    print(f"INFO: DOI {doi_to_retry} [RETRY] - PubMed Central (direct) failed. Attempting PubMed Central (Selenium method)...")
                    pmc_selenium_pdf_content_retry, pmc_selenium_status_msg_retry = download_with_selenium_pmc(driver, doi_to_retry, effective_title_for_retry)
                    if pmc_selenium_pdf_content_retry:
                        pdf_content_retry = pmc_selenium_pdf_content_retry
                        retry_successful_this_doi = True
                        successful_mirror_for_retry = pmc_selenium_status_msg_retry
                        overall_retry_status = pmc_selenium_status_msg_retry
                        mirror_attempts_details_for_retry.append(("PubMed Central (Selenium Retry)", "OBTENIDO", pmc_selenium_status_msg_retry))
                        temp_detailed_status_for_retry_log = f"Success_RETRY_PMC_Selenium_{pmc_selenium_status_msg_retry}"
                        temp_failure_reason_for_retry_log = ""
                    else:
                        mirror_attempts_details_for_retry.append(("PubMed Central (Selenium Retry)", "FALLO", pmc_selenium_status_msg_retry))
                        temp_failure_reason_for_retry_log = pmc_selenium_status_msg_retry
                        temp_detailed_status_for_retry_log = f"Failure_RETRY_PubMedCentral_Selenium_{pmc_selenium_status_msg_retry}"

                retry_end_time_actual_attempt = datetime.now()
                original_article_log_entry = next((log for log in all_articles_log if str(log.get('DOI', log.get('doi', ''))).strip() == doi_to_retry), None)

                if retry_successful_this_doi and pdf_content_retry:
                    successful_downloads += 1 # Increment before log
                    # overall_retry_status is already set if successful
                    if not overall_retry_status.startswith("OBTENIDO"):
                         overall_retry_status = "OBTENIDO"
                    # ... (update successful_articles_data, save PDF, add to ZIP as before) ...
                    articles_successfully_retried_ids.append(doi_to_retry)
                    original_data_for_success = {k: v for k, v in failed_article_entry.items() if k not in ['Failure_Reason', 'Detailed_Status', 'original_index']}; original_data_for_success['Successful_Mirror'] = successful_mirror_for_retry; successful_articles_data.append(original_data_for_success)
                    temp_dir_retry = "temp_scihub_pdfs"; 
                    if not os.path.exists(temp_dir_retry): os.makedirs(temp_dir_retry)
                    temp_pdf_path_retry = os.path.join(temp_dir_retry, f"temp_RETRY_{os.getpid()}_{retry_idx}_{pdf_filename_in_zip_retry}")
                    with open(temp_pdf_path_retry, 'wb') as f: f.write(pdf_content_retry)
                    try: total_downloaded_size_bytes += os.path.getsize(temp_pdf_path_retry)
                    except OSError as e: print(f"Advertencia: tamaño temp (reintento) {temp_pdf_path_retry}: {e}")
                    try:
                        with zipfile.ZipFile(zip_path, 'a', zipfile.ZIP_DEFLATED) as zf_append: zf_append.write(temp_pdf_path_retry, arcname=pdf_filename_in_zip_retry)
                    except Exception as e: print(f"Error CRÍTICO al agregar PDF (reintento) '{pdf_filename_in_zip_retry}' al ZIP: {e}")
                    temp_pdf_paths.append(temp_pdf_path_retry) 
                    if original_article_log_entry: # Update original log entry
                        original_article_log_entry['Detailed_Status'] = temp_detailed_status_for_retry_log if not ("Google Scholar" in successful_mirror_for_retry) else f"Success_RETRY_GoogleScholar_{successful_mirror_for_retry}"
                        original_article_log_entry['Failure_Reason'] = "" if retry_successful_this_doi else temp_failure_reason_for_retry_log
                        original_article_log_entry['Successful_Mirror'] = successful_mirror_for_retry
                        original_article_log_entry['End_Time'] = retry_end_time_actual_attempt.strftime("%Y-%m-%d %H:%M:%S")
                        original_start_dt = datetime.strptime(original_article_log_entry['Start_Time'], "%Y-%m-%d %H:%M:%S")
                        original_article_log_entry['Duration_Seconds'] = (retry_end_time_actual_attempt - original_start_dt).total_seconds()
                else:
                    # overall_retry_status is already FALTANTE by default or set by GS if it failed
                    if not overall_retry_status.startswith("FALLO"): # ensure it is if we are in this block.
                        overall_retry_status = "FALTANTE"
                    if original_article_log_entry: # Update original log entry with latest failure
                        original_article_log_entry['Detailed_Status'] = temp_detailed_status_for_retry_log
                        original_article_log_entry['Failure_Reason'] = temp_failure_reason_for_retry_log
                    for item in failed_articles_data: # Update reason in failed_articles_data for Excel
                        if str(item.get('DOI',item.get('doi',''))).strip() == doi_to_retry: item['Failure_Reason'] = temp_failure_reason_for_retry_log; item['Detailed_Status'] = temp_detailed_status_for_retry_log; break
                
                # Log for this retry attempt
                # Pass original_row_data from failed_article_entry for Revista, Fecha Pub, etc.
                format_and_log_article_status(failed_article_entry, doi_to_retry, effective_title_for_retry, current_article_num_for_log_retry, total_articles, successful_downloads, mirror_attempts_details_for_retry, overall_retry_status, user_inter_doi_delay, is_retry=True)

                # Separator after logging for the current article in the retry loop
                print_to_console("===============================================================================================", original_stdout)
                # if log_window and log_window.winfo_exists(): log_window.update_idletasks() # GUI logging disabled
                if retry_idx < len(temp_failed_articles_data_for_iteration) - 1: time.sleep(user_inter_doi_delay)
            
            if articles_successfully_retried_ids: failed_articles_data = [item for item in failed_articles_data if str(item.get('DOI', item.get('doi', ''))).strip() not in articles_successfully_retried_ids]

            # if log_window and log_window.winfo_exists(): log_window.update_idletasks() # GUI logging disabled
            failed_downloads_summary_list = [{'title': str(item.get('Title',item.get('title','N/A'))).strip(), 'doi': str(item.get('DOI',item.get('doi','N/A'))).strip(), 'reason': str(item.get('Failure_Reason','N/A')).strip()} for item in failed_articles_data]
            total_mb = total_downloaded_size_bytes / (1024 * 1024)
            summary_message = (f"Proceso completado.\n\nDescargas exitosas: {successful_downloads}\nDescargas fallidas: {len(failed_downloads_summary_list)}\n" f"Tamaño total PDFs: {total_mb:.2f} MB")
            if failed_downloads_summary_list:
                summary_message += "\n\nArtículos no descargados (post-reintentos):"
                for item in failed_downloads_summary_list:
                    summary_message += f"\n- Título: {item['title']}, DOI: {item['doi']}, Razón: {item['reason']}"
            print("\n" + "="*50); print(summary_message); print("="*50); 
            # if log_window and log_window.winfo_exists(): log_window.update_idletasks() # GUI logging disabled
            messagebox.showinfo("Resumen Descarga", summary_message) # Keep messagebox for summary
            generate_excel_report_prompt = messagebox.askyesno("Generar Reporte Excel", "¿Desea generar un reporte Excel detallado?")
            if generate_excel_report_prompt:
                excel_report_path_to_use = excel_report_path_config
                if not excel_report_path_to_use : 
                    excel_report_path_to_use = filedialog.asksaveasfilename(title="Guardar Reporte Excel como...", defaultextension=".xlsx", initialfile="SciHub_Descarga_Reporte.xlsx", filetypes=(("Archivos Excel", "*.xlsx"),("Todos los archivos", "*.*")))
                elif not messagebox.askyesno("Confirmar Ruta de Reporte", f"Se configuró guardar el reporte en:\n{excel_report_path_config}\n\n¿Usar esta ruta?"): 
                        excel_report_path_to_use = filedialog.asksaveasfilename(title="Guardar Reporte Excel en ruta alternativa...", defaultextension=".xlsx", initialfile="SciHub_Descarga_Reporte_Alternativo.xlsx", filetypes=(("Archivos Excel", "*.xlsx"), ("Todos los archivos", "*.*")))
                if excel_report_path_to_use:
                    print(f"Generando reporte Excel en: {excel_report_path_to_use}")
                    # if log_window and log_window.winfo_exists(): log_window.update_idletasks() # GUI logging disabled
                    try:
                        ob_cols = ['DOI','Title','Successful_Mirror'] + [c for c in original_input_columns if c not in ['DOI','Title','Successful_Mirror']] + ['SciHub_Link']
                        fa_cols = ['DOI','Title'] + [c for c in original_input_columns if c not in ['DOI','Title','Failure_Reason','Detailed_Status']] + ['Failure_Reason','Detailed_Status','SciHub_Link']
                        ti_cols = ['DOI','Title'] + [c for c in original_input_columns if c not in ['DOI','Title','Successful_Mirror','Start_Time','End_Time','Duration_Seconds','Detailed_Status','Failure_Reason']] + ['Successful_Mirror','Start_Time','End_Time','Duration_Seconds','Detailed_Status','Failure_Reason','SciHub_Link']
                        def create_ordered_df(data, cols):
                            df_ = pd.DataFrame(data); df_['SciHub_Link'] = df_.apply(lambda r: f"{sci_hub_base_url_for_report}{r.get('DOI',r.get('doi',''))}" if pd.notna(r.get('DOI',r.get('doi',''))) else '', axis=1)
                            for c in cols: 
                                if c not in df_.columns: df_[c] = pd.NA
                            # Ensure all original_row_data keys also get included if not in standard cols
                            all_data_keys = set()
                            if data: all_data_keys.update(k for item in data for k in item.keys())
                            final_cols_ordered = [c for c in cols if c in all_data_keys or c == 'SciHub_Link'] # Start with preferred that exist or are added
                            final_cols_ordered.extend([k for k in all_data_keys if k not in final_cols_ordered and k != 'SciHub_Link']) # Add remaining data keys
                            return df_.reindex(columns=final_cols_ordered)

                        df_obtenidos = create_ordered_df(successful_articles_data, ob_cols)
                        df_fallidos = create_ordered_df(failed_articles_data, fa_cols)
                        df_tiempos = create_ordered_df(all_articles_log, ti_cols)
                        with pd.ExcelWriter(excel_report_path_to_use, engine='openpyxl') as writer:
                            df_obtenidos.to_excel(writer, sheet_name='Obtenidos', index=False); df_fallidos.to_excel(writer, sheet_name='Fallidos', index=False); df_tiempos.to_excel(writer, sheet_name='Tiempos', index=False)
                        messagebox.showinfo("Reporte Excel", f"Reporte Excel guardado: {excel_report_path_to_use}"); print(f"Reporte Excel guardado: {excel_report_path_to_use}")
                    except Exception as e: messagebox.showerror("Error Guardando Excel", f"No se pudo guardar reporte Excel: {e}"); print(f"Error Excel: {e}")
                else: print("Generación de reporte Excel omitida (ruta no proporcionada).")
            else: print("Generación de reporte Excel omitida por usuario.")

        elif zip_creation_or_main_loop_error: print("Proceso interrumpido por error crítico inicial. No se generará resumen ni Excel.")
    
    finally: 
        # WebDriver will be quit here
        if driver:
            print("Cerrando WebDriver de Selenium...")
            try:
                driver.quit()
                print("WebDriver de Selenium cerrado correctamente.")
            except Exception as e:
                print(f"Error al cerrar WebDriver de Selenium: {e}")
        # restored_to_original_console = False # Not needed, stdout not redirected
        # Ensure stdout is the original, in case it was somehow still a TextRedirector
        # (though it shouldn't be if TextRedirector class and its usage are commented out)
        if hasattr(original_stdout, 'write'): # Check if original_stdout was captured
             if sys.stdout != original_stdout:
                 sys.stdout = original_stdout
                 print("\n--- stdout restaurado a la consola original durante la limpieza final ---", file=original_stdout)

        print("\n--- Limpieza Final de Archivos Temporales ---", file=original_stdout if hasattr(original_stdout, 'write') else sys.__stdout__)

        for temp_path in temp_pdf_paths:
            try:
                os.remove(temp_path)
                print(f"Eliminado temp: {temp_path}", file=original_stdout if hasattr(original_stdout, 'write') else sys.__stdout__)
            except OSError as e:
                print(f"Error eliminando temp {temp_path}: {e}", file=original_stdout if hasattr(original_stdout, 'write') else sys.__stdout__)

        temp_dir_to_check = "temp_scihub_pdfs"
        if os.path.exists(temp_dir_to_check) and not os.listdir(temp_dir_to_check):
            try: os.rmdir(temp_dir_to_check); print(f"Eliminado dir temp: {temp_dir_to_check}", file=original_stdout if hasattr(original_stdout, 'write') else sys.__stdout__)
            except OSError as e: print(f"Error eliminando dir temp {temp_dir_to_check}: {e}", file=original_stdout if hasattr(original_stdout, 'write') else sys.__stdout__)

        print("--- Limpieza Finalizada ---", file=original_stdout if hasattr(original_stdout, 'write') else sys.__stdout__)

        # if log_window: # GUI logging disabled
        #     try:
        #         if log_window.winfo_exists(): log_window.destroy()
        #     except tk.TclError: pass

        # Destroy the main hidden Tkinter window if it exists
        if 'root' in locals() and isinstance(root, tk.Tk):
            try:
                root.destroy()
            except tk.TclError:
                pass


if __name__ == "__main__":
    # Store the true original stdout at the very beginning if not already done for the main function
    # This is a fallback if download_pdfs_from_file itself had an issue before original_stdout was set.
    # However, original_stdout inside download_pdfs_from_file should be sys.__stdout__ if not redirected.
    # The check `isinstance(sys.stdout, TextRedirector)` is no longer relevant if TextRedirector is removed.

    # The critical part is ensuring sys.stdout is sys.__stdout__ if it got changed.
    # Since TextRedirector is commented out, sys.stdout should not be an instance of it.
    # The `finally` block in `download_pdfs_from_file` now tries to restore original_stdout.
    # A final check here could be made, but might be redundant if `download_pdfs_from_file` handles it.

    # Safest is to assume download_pdfs_from_file cleans up its own stdout if it changed it.
    # If TextRedirector and its instantiation are fully removed, sys.stdout should remain as console.
    download_pdfs_from_file()
    # No need to check for TextRedirector instance if it's removed.
    # if isinstance(sys.stdout, TextRedirector) and hasattr(sys.stdout, 'original_stdout'):
    #    sys.stdout = sys.stdout.original_stdout # Restore if it was a TextRedirector
    # elif isinstance(sys.stdout, TextRedirector): # Fallback if original_stdout attribute missing
    #    sys.stdout = sys.__stdout__

    print("\nScript finalizado.")
