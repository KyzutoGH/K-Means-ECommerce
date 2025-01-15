# Import library yang diperlukan
import json
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import seaborn as sns
from datetime import datetime
from collections import Counter

def load_data(json_data):
    """
    Fungsi untuk memuat dan membersihkan data dari JSON
    Args:
        json_data: Data JSON yang akan diproses
    Returns:
        DataFrame pandas yang sudah dibersihkan
    """
    df = pd.DataFrame(json_data)
    # Membersihkan kolom Omset dengan menghapus koma dan mengkonversi ke float
    df['Omset'] = df['Omset'].str.replace(',', '').astype(float)
    return df

def calculate_distances(omset, centroids):
    """
    Menghitung jarak antara omset dengan setiap centroid
    Args:
        omset: Nilai omset yang akan dihitung jaraknya
        centroids: List nilai centroid
    Returns:
        List jarak ke setiap centroid
    """
    return [abs(omset - centroid) for centroid in centroids]

def assign_cluster(distances):
    """
    Menentukan cluster berdasarkan jarak terdekat
    Args:
        distances: List jarak ke setiap centroid
    Returns:
        Nomor cluster (1-3)
    """
    return distances.index(min(distances)) + 1

def analyze_cluster_characteristics(results_df):
    """
    Menganalisis karakteristik setiap cluster
    Args:
        results_df: DataFrame hasil clustering
    Returns:
        Dictionary berisi analisis per cluster
    """
    cluster_analysis = {}
    
    for cluster in [1, 2, 3]:
        # Filter data untuk cluster tertentu
        cluster_data = results_df[results_df['Calculated Cluster'] == cluster]
        
        # Hitung rata-rata omset
        avg_omset = cluster_data['Omset'].mean()
        
        # Identifikasi produk dominan (3 teratas)
        product_counts = Counter(cluster_data['nama Produk'])
        top_products = product_counts.most_common(3)
        
        # Tentukan karakteristik berdasarkan cluster
        if cluster == 1:
            characteristics = "Toko dengan penjualan rendah atau baru memulai"
        elif cluster == 2:
            characteristics = "Toko dengan stabilitas penjualan menengah"
        else:
            characteristics = "Toko dengan performa penjualan tinggi"
            
        # Simpan hasil analisis
        cluster_analysis[cluster] = {
            'avg_omset': avg_omset,
            'characteristics': characteristics,
            'dominant_products': [prod[0] for prod in top_products]
        }
    
    return cluster_analysis

def create_excel_report(results_df, centroids, cluster_analysis):
    """
    Membuat laporan Excel lengkap
    Args:
        results_df: DataFrame hasil analisis
        centroids: List nilai centroid
        cluster_analysis: Dictionary hasil analisis cluster
    Returns:
        Nama file Excel yang dihasilkan
    """
    # Buat file Excel dengan timestamp
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    excel_filename = f'clustering_analysis_{timestamp}.xlsx'
    writer = pd.ExcelWriter(excel_filename, engine='xlsxwriter')
    
    # Set format untuk Excel
    workbook = writer.book
    header_format = workbook.add_format({
        'bold': True,
        'bg_color': '#D9E1F2',
        'border': 1
    })
    number_format = workbook.add_format({
        'num_format': '#,##0.00',
        'border': 1
    })
    text_format = workbook.add_format({
        'border': 1
    })
    
    # 1. Buat sheet Detail Results
    detailed_results = results_df[['Data id', 'Nama Toko', 'nama Produk', 'Omset', 
                                 'Calculated Cluster', 'Existing Cluster']]
    detailed_results.to_excel(writer, sheet_name='Detailed Results', index=False)
    
    # Format sheet Detail Results
    worksheet = writer.sheets['Detailed Results']
    worksheet.set_column('A:C', 20, text_format)
    worksheet.set_column('D:D', 15, number_format)
    worksheet.set_column('E:F', 12, text_format)
    
    for col_num, value in enumerate(detailed_results.columns.values):
        worksheet.write(0, col_num, value, header_format)
    
    # 2. Buat sheet Summary Statistics
    summary_sheet = writer.book.add_worksheet('Summary Statistics')
    
    # Hitung distribusi cluster
    calc_dist = results_df['Calculated Cluster'].value_counts().sort_index()
    exist_dist = results_df['Existing Cluster'].value_counts().sort_index()
    
    # Siapkan data ringkasan
    summary_data = {
        'Metric': [
            'Total Records',
            'Matching Clusters',
            'Match Percentage',
            'Cluster 1 Count (Calculated)',
            'Cluster 2 Count (Calculated)',
            'Cluster 3 Count (Calculated)',
            'Cluster 1 Count (Existing)',
            'Cluster 2 Count (Existing)',
            'Cluster 3 Count (Existing)',
        ],
        'Value': [
            len(results_df),
            sum(results_df['Calculated Cluster'] == results_df['Existing Cluster']),
            (sum(results_df['Calculated Cluster'] == results_df['Existing Cluster']) / len(results_df)) * 100,
            calc_dist.get(1, 0),
            calc_dist.get(2, 0),
            calc_dist.get(3, 0),
            exist_dist.get(1, 0),
            exist_dist.get(2, 0),
            exist_dist.get(3, 0),
        ]
    }
    
    # Tulis data ringkasan
    summary_sheet.write('A1', 'Summary Statistics', header_format)
    for i, (metric, value) in enumerate(zip(summary_data['Metric'], summary_data['Value'])):
        summary_sheet.write(i + 1, 0, metric, text_format)
        if isinstance(value, float):
            summary_sheet.write(i + 1, 1, value, number_format)
        else:
            summary_sheet.write(i + 1, 1, value, text_format)
    
    # Tambahkan karakteristik cluster
    row_offset = len(summary_data['Metric']) + 2
    summary_sheet.write(row_offset, 0, 'Cluster Characteristics', header_format)
    summary_sheet.write(row_offset, 1, '', header_format)
    
    # Tulis informasi setiap cluster
    for cluster in [1, 2, 3]:
        cluster_info = cluster_analysis[cluster]
        base_row = row_offset + (cluster - 1) * 4 + 1
        
        summary_sheet.write(base_row, 0, f'Cluster {cluster}', text_format)
        summary_sheet.write(base_row, 1, cluster_info['characteristics'], text_format)
        summary_sheet.write(base_row + 1, 0, 'Average Omset', text_format)
        summary_sheet.write(base_row + 1, 1, cluster_info['avg_omset'], number_format)
        summary_sheet.write(base_row + 2, 0, 'Dominant Products', text_format)
        summary_sheet.write(base_row + 2, 1, ', '.join(cluster_info['dominant_products']), text_format)
    
    summary_sheet.set_column('A:A', 30)
    summary_sheet.set_column('B:B', 50)
    
    # 3. Buat sheet Mismatches
    mismatches = results_df[results_df['Calculated Cluster'] != results_df['Existing Cluster']]
    mismatches = mismatches[['Data id', 'Nama Toko', 'nama Produk', 'Omset', 
                            'Calculated Cluster', 'Existing Cluster']]
    mismatches.to_excel(writer, sheet_name='Mismatches', index=False)
    
    # Format sheet Mismatches
    worksheet = writer.sheets['Mismatches']
    worksheet.set_column('A:C', 20, text_format)
    worksheet.set_column('D:D', 15, number_format)
    worksheet.set_column('E:F', 12, text_format)
    
    for col_num, value in enumerate(mismatches.columns.values):
        worksheet.write(0, col_num, value, header_format)
    
    # 4. Buat sheet Centroids
    centroid_sheet = writer.book.add_worksheet('Centroids')
    centroid_sheet.write('A1', 'Cluster', header_format)
    centroid_sheet.write('B1', 'Centroid Value', header_format)
    
    for i, centroid in enumerate(centroids, 1):
        centroid_sheet.write(i, 0, f'Cluster {i}', text_format)
        centroid_sheet.write(i, 1, centroid, number_format)
    
    centroid_sheet.set_column('A:A', 15)
    centroid_sheet.set_column('B:B', 20)
    
    # Simpan file Excel
    writer.close()
    return excel_filename

def main():
    """
    Fungsi utama yang menjalankan seluruh proses analisis
    """
    # Baca file JSON
    with open('datasetnew.json', 'r') as file:
        json_data = json.load(file)
    
    # Konversi ke DataFrame
    df = load_data(json_data)
    
    # Tentukan centroid
    centroids = [424000.00, 915000.00, 689155580.85]
    
    # Proses clustering
    results = []
    for index, row in df.iterrows():
        omset = row['Omset']
        distances = calculate_distances(omset, centroids)
        assigned_cluster = assign_cluster(distances)
        
        # Tentukan existing cluster
        existing_cluster = None
        if row['Kluster 1'] == '1':
            existing_cluster = 1
        elif row['Kluster 2'] == '1':
            existing_cluster = 2
        elif row['Kluster 3'] == '1':
            existing_cluster = 3
        
        # Simpan hasil
        results.append({
            'Data id': row['Data id'],
            'Nama Toko': row['Nama Toko'],
            'nama Produk': row['nama Produk'],
            'Omset': omset,
            'Calculated Cluster': assigned_cluster,
            'Existing Cluster': existing_cluster,
            'Distances': distances
        })
    
    # Konversi hasil ke DataFrame
    results_df = pd.DataFrame(results)
    
    # Analisis karakteristik cluster
    cluster_analysis = analyze_cluster_characteristics(results_df)
    
    # Cetak hasil analisis ke console
    print("\nCluster Analysis Results:")
    for cluster, info in cluster_analysis.items():
        print(f"\nCluster {cluster}:")
        print(f"Rata-rata omset: Rp {info['avg_omset']:,.2f}")
        print(f"Karakteristik umum: {info['characteristics']}")
        print(f"Produk dominan: {', '.join(info['dominant_products'])}")
    
    # Buat laporan Excel
    excel_filename = create_excel_report(results_df, centroids, cluster_analysis)
    print(f"\nExcel report has been generated: {excel_filename}")

# Jalankan program jika dieksekusi langsung
if __name__ == "__main__":
    main()