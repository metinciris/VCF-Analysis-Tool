# VCF Analysis Tool

This project provides a tool to analyze multiple VCF files (Unfiltered_variants), identify shared mutations, and summarize the results with associated read depths and quality values in an Excel file. The tool is user-friendly and includes a detailed explanation of the results within the Excel output.

## Features
- Parse multiple VCF files and extract mutation information.
- Identify shared mutations across multiple samples.
- Include key details such as chromosome, position, reference base, alternate base, read depth, and quality.
- Save results in an organized Excel file with an additional explanation sheet.

## Requirements
- Python 3.6 or above
- The following Python modules:
  - `pandas`
  - `xlsxwriter`
  - `tkinter`

Install the required modules using pip:
```bash
pip install pandas xlsxwriter
```

## Usage

1. Clone the repository:
   ```bash
   git clone https://github.com/your-username/vcf-analysis-tool.git
   cd vcf-analysis-tool
   ```

2. Run the script:
   ```bash
   python vcf_analysis_tool.py
   ```

3. Select multiple VCF files through the file dialog.

4. Choose to either add more files or perform the analysis.

5. Save the results to an Excel file.

## Output Example
### Results Sheet
| Kromozom | Pozisyon | Referans Baz | Alternatif Baz | Sample1           | Sample2           | Sample3           |
|----------|----------|--------------|----------------|-------------------|-------------------|-------------------|
| 1        | 12345    | A            | T              | 444/5555 / 99.9   | 333/4444 / 98.5   |                   |
| 1        | 54321    | G            | C              |                   | 222/3333 / 85.2   | 111/2222 / 90.0   |

### Explanation Sheet
| Column           | Description                                          |
|------------------|------------------------------------------------------|
| Kromozom         | Chromosome where the mutation is located.            |
| Pozisyon         | Position of the mutation on the chromosome.          |
| Referans Baz     | The reference base (original DNA sequence).          |
| Alternatif Baz   | The alternate base (mutated DNA sequence).           |
| Sample Columns   | Read depth and quality for each sample (e.g., 2/288 / 13.06). |

## Example Input
- Multiple VCF files with format:
  ```
  #CHROM POS ID REF ALT QUAL FILTER INFO FORMAT Sample1
  1      12345 .  A   T   99.9 PASS   .    GT:AD    0/1:20,444
  ```

## Contribution
Feel free to fork the repository and submit pull requests for new features or bug fixes.

## License
This project is licensed under the MIT License.
