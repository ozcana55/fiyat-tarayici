import time
import io
import pandas as pd
import streamlit as st
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from webdriver_manager.chrome import ChromeDriverManager

# --- SAYFA AYARLARI ---
st.set_page_config(page_title="Fiyat Tarayıcı", page_icon="🔍", layout="centered")

st.title("🔍 Akakçe Fiyat Karşılaştırma Botu")
st.write("Aratmak istediğiniz ürünü girin, güncel fiyatları bulup Excel olarak indirelim.")

# --- SÜRÜCÜ HAZIRLAMA ---
def surucu_baslat():
    options = Options()
    options.add_argument("--headless")  # Arka planda çalışması için (Web sitelerinde ekran açılmaz)
    options.add_argument("--no-sandbox")
    options.add_argument("--disable-dev-shm-usage")
    options.add_argument("--disable-notifications")
    options.add_argument("--disable-popup-blocking")
    options.add_argument("user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36")
    
    return webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=options)

# --- VERİ ÇEKME FONKSİYONU ---
def akakce_arama(kelime):
    driver = surucu_baslat()
    arama_url = f"https://www.akakce.com/arama/?q={kelime.replace(' ', '+')}"
    
    try:
        driver.get(arama_url)
        time.sleep(3)
        driver.execute_script("window.scrollTo(0, 800);")
        time.sleep(2)
        
        urun_listesi = []
        urunler = driver.find_elements(By.CSS_SELECTOR, "ul#CPL > li, li.p_v8, li[data-pr]")

        for urun in urunler:
            try:
                # Başlık
                try:
                    baslik = urun.find_element(By.CSS_SELECTOR, "h3, span.pn_v8, a.pn_v8").text.strip()
                except:
                    baslik = urun.text.split("\n")[0].strip()

                # Link
                try:
                    link = urun.find_element(By.CSS_SELECTOR, "a").get_attribute("href")
                except:
                    link = "Link Bulunamadı"

                # Fiyat
                fiyat = ""
                fiyat_elementleri = urun.find_elements(By.CSS_SELECTOR, "span.pt_v8, span.pr_v8, div.pr_v8, em.pb_v8, span.pb_v8, b.pt_v8")
                for fe in fiyat_elementleri:
                    if fe.text.strip():
                        fiyat = fe.text.strip()
                        break
                
                if not fiyat:
                    for s in urun.text.split("\n"):
                        if "TL" in s or "tl" in s:
                            fiyat = s.strip()
                            break

                # Satıcı
                satici = "Akakçe Detayında"
                try:
                    satici_elem = urun.find_element(By.CSS_SELECTOR, "span.v_v8, span.v_b8, span.sb_v8, span.v_b")
                    satici = satici_elem.text.strip()
                except:
                    pass

                if baslik and fiyat:
                    urun_listesi.append({
                        "Aranan Kelime": kelime,
                        "Ürün / Model Adı": baslik,
                        "En Uygun Fiyat": fiyat.replace("\n", " "),
                        "Satıcı / Mağaza Bilgisi": satici,
                        "Ürün Linki": link
                    })
            except Exception:
                continue

        return urun_listesi
    finally:
        driver.quit()

# --- STREAMLIT ARAYÜZÜ ---
aranan_urun = st.text_input("Aratılacak Ürün:", placeholder="Örn: Playstation 5, Toshiba 65 TV, Bebek Arabası")

if st.button("Fiyatları Getir", type="primary"):
    if aranan_urun.strip():
        with st.spinner("Fiyatlar taranıyor, lütfen bekleyin..."):
            veri = akakce_arama(aranan_urun)
            
            if veri:
                df = pd.DataFrame(veri)
                st.success(f"Başarılı! **{len(veri)}** adet ürün bulundu.")
                
                # Tabloyu Ekranda Göster
                st.dataframe(df[["Ürün / Model Adı", "En Uygun Fiyat", "Satıcı / Mağaza Bilgisi"]], use_container_width=True)
                
                # Excel Dosyasını Bellekte Oluşturup İndirme Butonu Koyma
                buffer = io.BytesIO()
                with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
                    df.to_excel(writer, index=False)
                
                temiz_dosya_adi = "".join(c for c in aranan_urun if c.isalnum() or c in (' ', '_')).rstrip().replace(' ', '_')
                
                st.download_button(
                    label="📥 Excel Dosyasını İndir",
                    data=buffer.getvalue(),
                    file_name=f"{temiz_dosya_adi}_Fiyatlari.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
            else:
                st.error("Ürün bulunamadı veya veri çekilemedi.")
    else:
        st.warning("Lütfen bir ürün adı girin!")