import pandas as pd

# İncelenecek dosya isimleri
file_names = ['YBCG', 'YCFS', 'YLHI', 'YSCB', 'YSDU', 'YSNF', 'YSRI', 'YSSY', 'YSTW', 'YWLM']
folder_path = r"C:\Users\aleyna\OneDrive\Masaüstü\atmos\METAR_SPECI\METAR_SPECI"  # Dosyaların bulunduğu klasör

# İncelenecek yıllar ve hava olayları
years = list(range(2014, 2025))
weather_events = [
    '+RA', '+DZ', '+SHRA', '+SHGR', '+SHGS', '+TSRA', '+TSGR', '+TSGS', '+SHRADZ']

# Her dosya için işlemleri gerçekleştir
for file_name in file_names:
    # CSV dosyasının tam yolunu oluştur
    file_path = f"{folder_path}\\{file_name}.csv"
    
    # CSV dosyasını oku 
    df = pd.read_csv(file_path)

    # Tarih sütununu datetime formatına çevir (Tarih sütununun adı 'valid')
    df['valid'] = pd.to_datetime(df['valid'])

    # Yıllık ve toplam sayıları tutacak sözlükler
    annual_counts = {event: {year: 0 for year in years} for event in weather_events}
    total_counts = {event: 0 for event in weather_events}

    # Yıllık hava olayı sayısını bul
    for event in weather_events:
        for year in years:
            count = df[(df['valid'].dt.year == year) & (df['metar'].str.contains(event.replace('+', r'\+')))].shape[0]
            annual_counts[event][year] = count
            total_counts[event] += count

    # Yıllık sayıları yazdır ve kaydetmek için bir listeye ekle
    output_lines = []
    for event in weather_events:
        output_lines.append(f"Hava Olayı: {event}")
        for year in years:
            line = f"{year}: {annual_counts[event][year]} tane"
            output_lines.append(line)

    # Toplam sayıları yazdır ve listeye ekle
    output_lines.append("\nToplam Sayılar:")
    for event in weather_events:
        line = f"{event}: {total_counts[event]} tane"
        output_lines.append(line)

    # Yıllık sayıları ve toplamları içeren bir DataFrame oluştur
    results = pd.DataFrame(annual_counts).T
    results['Total'] = results.sum(axis=1)

    # Sütunları başlıklara göre yeniden adlandırma
    results.columns = [str(year) for year in years] + ['Total']

    # ExcelWriter kullanarak iki farklı sayfaya yazma
    output_excel_file = f"{file_name}_hava_olayi_sayilari.xlsx"
    with pd.ExcelWriter(output_excel_file, engine='xlsxwriter') as writer:
        # Listeyi bir sayfaya yaz
        pd.DataFrame(output_lines, columns=['Data']).to_excel(writer, sheet_name='Liste', index=False)

        # Tabloyu diğer sayfaya yaz
        results.index.name = 'Weather Event'
        results.to_excel(writer, sheet_name='Tablo')

    print(f"Sonuçlar '{output_excel_file}' dosyasına iki farklı sayfada kaydedildi.")
