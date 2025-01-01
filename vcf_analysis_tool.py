import tkinter as tk
from tkinter import filedialog, messagebox
import pandas as pd

def parse_vcf_file(file_path):
    """Parse a VCF file to extract relevant mutation information."""
    data = []
    with open(file_path, 'r') as file:
        for line in file:
            if not line.startswith("#"):
                parts = line.strip().split('\t')
                chrom = parts[0]
                pos = parts[1]
                ref = parts[3]
                alt = parts[4]
                qual = parts[5]
                format_fields = parts[8].split(':')
                sample_data = parts[9].split(':')

                if "CLCAD2" in format_fields:
                    clcad2_index = format_fields.index("CLCAD2")
                    depths = sample_data[clcad2_index].split(',')
                    ref_depth = int(depths[0])
                    alt_depth = int(depths[1]) if len(depths) > 1 else 0
                else:
                    ref_depth, alt_depth = 0, 0

                mutation = f"{chrom}:{pos}"  # Simplified mutation identifier
                data.append((chrom, pos, ref, alt, mutation, alt_depth, ref_depth + alt_depth, qual))

    return pd.DataFrame(data, columns=['Kromozom', 'Pozisyon', 'Referans Baz', 'Alternatif Baz', 'Mutasyon', 'Alt Okuma', 'Toplam Okuma', 'Kalite'])

def create_overlap_table(vcf_files):
    """Create an overlap table from selected VCF files."""
    combined_data = {}

    for file in vcf_files:
        file_data = parse_vcf_file(file)
        simplified_name = file.split('/')[-1].split('\\')[-1]  # Simplify file name
        combined_data[simplified_name] = file_data

    mutations = set()
    for data in combined_data.values():
        mutations.update(data['Mutasyon'])

    result_table = pd.DataFrame({'Mutasyon': list(mutations)})

    for file, data in combined_data.items():
        file_dict = {
            row['Mutasyon']: f"{row['Alt Okuma']}/{row['Toplam Okuma']} / {row['Kalite']}"
            for _, row in data.iterrows()
        }
        result_table[file] = result_table['Mutasyon'].map(file_dict).fillna('')

    # Add detailed mutation columns directly from the parsed data
    mutation_details = pd.concat(combined_data.values(), ignore_index=True)
    mutation_details = mutation_details.drop_duplicates(subset=['Mutasyon'])
    result_table = result_table.merge(mutation_details[['Mutasyon', 'Kromozom', 'Pozisyon', 'Referans Baz', 'Alternatif Baz']], how='left', on='Mutasyon')
    result_table = result_table[['Kromozom', 'Pozisyon', 'Referans Baz', 'Alternatif Baz'] + [col for col in result_table.columns if col not in ['Kromozom', 'Pozisyon', 'Referans Baz', 'Alternatif Baz', 'Mutasyon']]]

    return result_table

def save_to_excel_with_info(df, output_path):
    """Save the result table to an Excel file with an additional sheet for explanation."""
    with pd.ExcelWriter(output_path, engine='xlsxwriter') as writer:
        df.to_excel(writer, sheet_name='Results', index=False)

        # Add explanation sheet
        explanation = {
            "Column": ["Kromozom", "Pozisyon", "Referans Baz", "Alternatif Baz", "Hasta Dosyası"],
            "Description": [
                "Mutasyonun bulunduğu kromozom bilgisi",
                "Mutasyonun genom üzerindeki pozisyonu",
                "Mutasyonun referans baz bilgisi (orijinal dizi)",
                "Mutasyon sonrası alternatif baz bilgisi (değişmiş dizi)",
                "Her bir hasta dosyası için Alt Okuma / Toplam Okuma ve Kalite bilgisi"
            ]
        }
        explanation_df = pd.DataFrame(explanation)
        explanation_df.to_excel(writer, sheet_name='Explanation', index=False)

def main():
    root = tk.Tk()
    root.withdraw()

    # Step 1: Select files
    vcf_files = filedialog.askopenfilenames(title="Select VCF Files", filetypes=[("VCF Files", "*.vcf")])
    if not vcf_files:
        print("No files selected.")
        return

    # Step 2: Ask for operation
    operation = messagebox.askquestion("Operation", "Do you want to add new files or perform analysis?", icon='question')

    if operation == 'yes':
        new_files = filedialog.askopenfilenames(title="Select Additional VCF Files", filetypes=[("VCF Files", "*.vcf")])
        if new_files:
            vcf_files = list(vcf_files) + list(new_files)
        else:
            print("No additional files selected.")

    # Step 3: Perform analysis
    overlap_table = create_overlap_table(vcf_files)

    # Step 4: Save the results
    output_path = filedialog.asksaveasfilename(title="Save As", defaultextension=".xlsx", filetypes=[("Excel Files", "*.xlsx")])
    if output_path:
        save_to_excel_with_info(overlap_table, output_path)
        print(f"Results saved to {output_path}")

if __name__ == "__main__":
    main()
