# sp5-logsheet
control room log sheet modifiication

## VBA code optimization (practical)
This repo contains workbook macros inside `SP5_Production_v1.xlsm`.

### 1) Generate module-wise optimized code directly from workbook
```bash
python tools/export_and_optimize_vba.py
```
This command creates `optimized_vba/*.bas` files (module-wise cleaned output) directly from `xl/vbaProject.bin`.

### 2) Stream extraction (analysis)
```bash
python tools/extract_vba.py
```
This extracts VBA streams into raw/decompressed artifacts in `extracted_vba/`.

### 3) Optimize any exported `.bas/.cls/.frm`
```bash
python tools/optimize_vba_module.py path/to/module.bas
```
Applies safe formatting cleanup (`Option Explicit`, spacing normalization, blank-line cleanup, and declaration spacing).
