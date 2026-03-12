# Well-being Index Calculator — PyQt6 Desktop Application

## Overview

โปรแกรม Desktop สำหรับคำนวณ **Well-being Index** จากไฟล์ SPSS (.sav) โดยอัตโนมัติ ครอบคลุมตั้งแต่การโหลดข้อมูล, คำนวณ Index, ไปจนถึง Export ผลลัพธ์เป็น Excel — ใช้ PyQt6 สำหรับ GUI ที่สวยงาม ทันสมัย

---

## 1. Functional Requirements

### 1.1 Input
- รับไฟล์ **SPSS (.sav)** ผ่าน file dialog หรือ drag-and-drop
- ไฟล์ต้องมีตัวแปร Q1, Q2 (satisfaction items ของแต่ละ Dimension), และ demographic variables (Country, Generation, Gender)
- ใช้ library `pyreadstat` สำหรับอ่านไฟล์ SPSS

### 1.2 Processing Pipeline (6 ขั้นตอน)

โปรแกรมต้องทำงานตาม pipeline นี้อย่างครบถ้วน:

#### Step 1 — โครงสร้างแบบสอบถาม
แบบสอบถามแบ่งเป็น **6 Dimensions** โดยแต่ละ Dimension ประกอบด้วย sub-items ดังนี้:

| Dimension | Prefix | Sub-items | Label |
|---|---|---|---|
| Dim1 | PWQ2 | #1 to #5 | Physical well-being |
| Dim2 | MWQ2 | #1 to #4 | Mental |
| Dim3 | SWQ2 | #1 to #8 | Social |
| Dim4 | FWQ2 | #1 to #10 | Financial |
| Dim5 | EWQ2 | #1 to #7 | Environmental |
| Dim6 | PSWQ2 | #1 to #6 | Purpose & spiritual |

นอกจากนี้มี **Q1** (overall feeling, 7-point scale จาก positive→negative)

**สำคัญ**: ชื่อตัวแปรในไฟล์ SPSS อาจไม่ตรงกับที่ระบุข้างต้นทุกครั้ง — โปรแกรมต้องมี **Variable Mapping UI** ให้ผู้ใช้ map ตัวแปรจากไฟล์ SPSS กับ Dimension ที่ต้องการ หรือใช้ auto-detect จาก variable labels

#### Step 2 — เตรียมตัวแปร

**2.1 Recode Q1** — กลับ scale ให้ positive = ค่ามาก:
```
Original → Recoded (NQ1)
1 → 7
2 → 6
3 → 5
4 → 4
5 → 3
6 → 2
7 → 1
```
สูตร: `NQ1 = 8 - Q1`

**2.2 คำนวณ Dimension means** — สร้างตัวแปร Dim1–Dim6 โดยหาค่าเฉลี่ยของ sub-items ในแต่ละ Dimension:
```python
Dim1 = mean(PWQ2_1, PWQ2_2, PWQ2_3, PWQ2_4, PWQ2_5)
Dim2 = mean(MWQ2_1, MWQ2_2, MWQ2_3, MWQ2_4)
Dim3 = mean(SWQ2_1, SWQ2_2, ..., SWQ2_8)
Dim4 = mean(FWQ2_1, FWQ2_2, ..., FWQ2_10)
Dim5 = mean(EWQ2_1, EWQ2_2, ..., EWQ2_7)
Dim6 = mean(PSWQ2_1, PSWQ2_2, ..., PSWQ2_6)
```
- ใช้ `np.nanmean()` เพื่อจัดการ missing values

#### Step 3 — Factor Analysis (PCA)
- รัน **Principal Component Analysis (PCA)** บน Dim1–Dim6 (6 ตัวแปร)
- ใช้ **EQUAMAX rotation** (หรือ varimax ถ้า equamax ไม่มี)
- Extract factors เท่ากับจำนวนตัวแปร (6 factors)
- Save **factor scores** สำหรับแต่ละ respondent
- ใช้ library `factor_analyzer` หรือ `sklearn.decomposition.PCA` + manual rotation
- **วัตถุประสงค์**: ทำให้ตัวแปร 6 ตัวเป็นอิสระต่อกัน (orthogonal) เพื่อลด multicollinearity ก่อน regression

```python
from factor_analyzer import FactorAnalyzer

fa = FactorAnalyzer(n_factors=6, rotation='equamax', method='principal')
fa.fit(dim_data)
factor_scores = fa.transform(dim_data)
```

#### Step 4 — Multiple Regression
- Dependent variable: **NQ1** (recoded Q1)
- Independent variables: **Factor scores** (FAC1–FAC6) จาก Step 3
- Method: Enter (บังคับใส่ทุกตัว)
- ดึงค่า **Standardized Beta coefficients** ของแต่ละ Factor
- รายงาน R², Adjusted R², F-statistic, Sig., Collinearity (Tolerance, VIF)

```python
import statsmodels.api as sm

X = sm.add_constant(factor_scores)
model = sm.OLS(NQ1, X).fit()
betas = model.params[1:]  # ไม่รวม constant
std_betas = ...  # standardized coefficients
```

#### Step 5 — คำนวณน้ำหนัก (Weight)
- เอาค่า **|Standardized Beta|** ของแต่ละ Dimension
- คำนวณ **proportional weight**: `weight_i = |beta_i| / sum(|beta_all|)`
- ผลรวม weights ต้อง = 1.0 (100%)

ตัวอย่างผลลัพธ์ (จากข้อมูลจริง):
| Dimension | Std. Beta | Weight |
|---|---|---|
| Physical | 0.2122 | 14.27% |
| Mental | 0.2956 | 19.87% |
| Social | 0.2702 | 18.16% |
| Financial | 0.2931 | 19.70% |
| Environmental | 0.1639 | 11.02% |
| Purpose & Spiritual | 0.2527 | 16.98% |

#### Step 6 — คำนวณ Well-being Index
```
Index = Σ (Dim_i × Weight_i)  สำหรับ i = 1..6
```
- คำนวณ Index สำหรับแต่ละ respondent
- จากนั้นคำนวณ **mean Index** สำหรับแต่ละ subgroup

### 1.3 Subgroup Analysis (Breakdown)
คำนวณผลลัพธ์แยกตาม subgroup ต่อไปนี้ (ทำซ้ำ Step 3–6 สำหรับแต่ละ subgroup):

**Country-level**: Total, Japan, China, Thailand, Vietnam
**Country × Generation**: Japan Gen X, Japan Gen Y, Japan Gen Z, China Gen X, ... (4×3 = 12 groups)
**Generation-level**: Total Gen X, Total Gen Y, Total Gen Z
**Gender-level**: Total Male, Total Female
**Country × Gender**: Japan Male, Japan Female, China Male, ... (4×2 = 8 groups)

**สำคัญมาก**: น้ำหนัก (weights) ต้อง **คำนวณแยกสำหรับแต่ละ subgroup** — ไม่ใช่ใช้ weight จาก Total ไปใช้กับทุกกลุ่ม เพราะแต่ละกลุ่มมี factor structure และ regression coefficients ที่แตกต่างกัน

### 1.4 Output — Excel File (3 Sheets)

#### Sheet 1: "Correlations_Q2" — Pearson Correlation ระหว่าง Index กับ Q2 sub-items
- **Rows**: แต่ละ sub-item (PWQ2#1, PWQ2#2, ..., PSWQ2#6) — รวม 40 items
- **Columns**: แต่ละ subgroup (Total, Japan, China, Thailand, Vietnam, Japan Gen X, ...)
- แต่ละ sub-item แสดง 3 แถว:
  1. **Pearson Correlation** coefficient (r)
  2. **Sig. (2-tailed)** p-value
  3. **N** (sample size)

#### Sheet 2: "Correlations_Q3" — Pearson Correlation ระหว่าง Index กับ Q3 behavioral items (MA)
- โครงสร้างเหมือน Sheet 1 แต่ใช้ Q3 items (PWQ3, MWQ4, SWQ3, FWQ3, EWQ3, PSWQ3)
- Q3 items เป็น **Multiple Answer (MA)** ดังนั้นค่าเป็น 0/1
- Columns เพิ่ม: Gen X, Gen Y, Gen Z (รวม generation ข้ามประเทศ)

#### Sheet 3: "Index" — สรุป Index Score และ Weights
- **Rows**: แต่ละ subgroup (Total, Japan, China, Thailand, Vietnam, Gen X, Gen Y, Gen Z, Japan Gen X, ..., Total Male, Total Female, Japan Male, ...)
- **Columns**:
  - `Index` — ค่า Well-being Index (weighted mean)
  - `score (1-7)`: Physical, Mental, Social, Financial, Environmental, Purpose & spiritual — ค่าเฉลี่ย Dim1–Dim6 (ก่อน weight)
  - `Weight`: Physical, Mental, Social, Financial, Environmental, Purpose & spiritual — น้ำหนักของแต่ละ Dimension (จาก Step 5)

---

## 2. Technical Architecture

### 2.1 Tech Stack
```
Python 3.11+
├── PyQt6              — GUI framework
├── pyreadstat          — Read SPSS .sav files
├── pandas              — Data manipulation
├── numpy               — Numerical computation
├── factor_analyzer     — Factor Analysis (PCA + rotation)
├── statsmodels         — Multiple Regression
├── scipy.stats         — Pearson correlation, p-values
├── openpyxl            — Write Excel output
└── pyinstaller         — Build standalone .exe (optional)
```

### 2.2 Project Structure
```
wellbeing-index-calculator/
├── pyproject.toml
├── README.md
├── src/
│   └── wellbeing/
│       ├── __init__.py
│       ├── main.py                 # Entry point
│       ├── config.py               # App settings, constants
│       ├── models/
│       │   ├── __init__.py
│       │   ├── dimensions.py       # Dimension definitions, variable mappings
│       │   └── subgroups.py        # Subgroup definitions (Country, Gen, Gender)
│       ├── services/
│       │   ├── __init__.py
│       │   ├── data_loader.py      # SPSS file loading & variable detection
│       │   ├── preprocessor.py     # Recode Q1, compute Dim means
│       │   ├── factor_service.py   # Factor Analysis (PCA + rotation)
│       │   ├── regression_service.py  # Multiple Regression
│       │   ├── index_calculator.py # Weight calculation & Index computation
│       │   ├── correlation_service.py # Pearson correlations for output
│       │   └── export_service.py   # Excel export (3 sheets)
│       ├── gui/
│       │   ├── __init__.py
│       │   ├── main_window.py      # Main window with navigation
│       │   ├── file_loader_widget.py  # File selection & variable mapping
│       │   ├── progress_widget.py  # Processing progress with logs
│       │   ├── results_widget.py   # Results display (tables, charts)
│       │   ├── styles.py           # QSS stylesheet (modern look)
│       │   └── workers.py          # QThread workers for heavy processing
│       └── utils/
│           ├── __init__.py
│           └── stats_helpers.py    # Statistical computation helpers
├── tests/
│   ├── conftest.py
│   ├── test_preprocessor.py
│   ├── test_factor_service.py
│   ├── test_regression_service.py
│   └── test_index_calculator.py
└── scripts/
    └── build_exe.py               # PyInstaller build script
```

### 2.3 Class Design

```python
# models/dimensions.py
from dataclasses import dataclass, field

@dataclass
class DimensionDef:
    """Definition of a single Well-being Dimension."""
    name: str                    # e.g., "Physical well-being"
    short_name: str              # e.g., "Physical"
    prefix: str                  # e.g., "PWQ2"
    q2_variables: list[str]      # SPSS variable names for Q2 sub-items
    q3_prefix: str               # e.g., "PWQ3" for behavioral items
    q3_variables: list[str]      # SPSS variable names for Q3 items

@dataclass 
class SubgroupDef:
    """Definition of an analysis subgroup."""
    name: str                    # e.g., "Japan Gen X"
    filter_col: str              # column name for filtering
    filter_value: Any            # value to filter on
    # Or use multiple filters
    filters: dict[str, Any] = field(default_factory=dict)

@dataclass
class IndexResult:
    """Result for one subgroup."""
    subgroup_name: str
    index_value: float
    dim_means: dict[str, float]  # Dim name → mean score
    dim_weights: dict[str, float]  # Dim name → weight
    regression_stats: dict       # R², F, Sig, etc.
    n: int
```

### 2.4 Processing Pipeline (Service Layer)

```python
class IndexCalculationPipeline:
    """Orchestrates the full calculation pipeline."""
    
    def __init__(self, df: pd.DataFrame, dimensions: list[DimensionDef]):
        self.df = df
        self.dimensions = dimensions
    
    def run_for_subgroup(self, mask: pd.Series) -> IndexResult:
        """Run full pipeline for a single subgroup."""
        sub_df = self.df[mask].copy()
        
        # Step 2: Preprocess
        sub_df = self.preprocessor.recode_q1(sub_df)
        dim_df = self.preprocessor.compute_dim_means(sub_df)
        
        # Step 3: Factor Analysis
        factor_scores = self.factor_service.run_pca(dim_df)
        
        # Step 4: Regression
        nq1 = sub_df['NQ1']
        reg_result = self.regression_service.run(nq1, factor_scores)
        
        # Step 5: Compute weights
        weights = self.index_calculator.compute_weights(reg_result.std_betas)
        
        # Step 6: Compute Index
        index_values = self.index_calculator.compute_index(dim_df, weights)
        
        return IndexResult(...)
    
    def run_all_subgroups(self) -> list[IndexResult]:
        """Run pipeline for all predefined subgroups."""
        results = []
        for subgroup in self.get_all_subgroups():
            mask = self.build_mask(subgroup)
            # Minimum sample size check
            if mask.sum() >= 30:
                result = self.run_for_subgroup(mask)
                results.append(result)
        return results
```

---

## 3. GUI Design (PyQt6)

### 3.1 Layout — 3-Step Wizard Style

```
┌─────────────────────────────────────────────────────────────┐
│  Well-being Index Calculator                          [—][×] │
├─────────────────────────────────────────────────────────────┤
│  ① Load Data  ──►  ② Configure  ──►  ③ Results             │
│  ─────────────────────────────────────────────────────       │
│                                                              │
│  ┌─────────────────────────────────────────────────────┐    │
│  │                                                      │    │
│  │              [Main Content Area]                      │    │
│  │                                                      │    │
│  │  Step 1: Drag & drop .sav file or click Browse       │    │
│  │  Step 2: Variable mapping + subgroup config          │    │
│  │  Step 3: Results table + charts + export             │    │
│  │                                                      │    │
│  └─────────────────────────────────────────────────────┘    │
│                                                              │
│  ┌─────────────────────────────────────────────────────┐    │
│  │  Processing Log (collapsible)                        │    │
│  │  [INFO] Loaded 1303 records from survey.sav          │    │
│  │  [INFO] Computing Dim1 (Physical): mean=5.19         │    │
│  │  [INFO] Factor Analysis complete. KMO=0.87           │    │
│  └─────────────────────────────────────────────────────┘    │
│                                                              │
│  ◄ Back                                    Next ► / Export  │
└─────────────────────────────────────────────────────────────┘
```

### 3.2 Step 1 — Load Data
- **Drop zone**: ลาก .sav ไฟล์มาวาง หรือกด Browse
- แสดง file info: ชื่อไฟล์, จำนวน records, จำนวน variables
- **Variable preview table**: แสดง variable name, label, type, n valid
- Auto-detect ตัวแปรจาก prefix (PWQ2, MWQ2, etc.)

### 3.3 Step 2 — Configure
- **Variable Mapping Panel**: ให้ user ยืนยัน/แก้ไข mapping
  - Q1 variable → dropdown เลือกจากตัวแปรในไฟล์
  - แต่ละ Dimension → multi-select ตัวแปร sub-items
  - Demographic variables: Country, Generation, Gender
- **Subgroup Configuration**: checkbox เลือก subgroups ที่ต้องการวิเคราะห์
- **Q3 Variable Mapping**: map behavioral items (ถ้ามี)
- **Advanced Settings** (collapsible):
  - Rotation method: Equamax / Varimax / Promax
  - Minimum sample size per subgroup (default: 30)
  - Missing value handling: listwise / pairwise

### 3.4 Step 3 — Results
- **Tab 1: Index Summary** — ตาราง Index score + Dim scores + Weights per subgroup
- **Tab 2: Regression Details** — R², Beta, Sig สำหรับแต่ละ subgroup
- **Tab 3: Correlations** — Pearson r between Index and each item
- **Export Button**: บันทึกเป็น Excel (3 sheets ตาม spec ข้างต้น)

### 3.5 Modern UI Styling

```python
# gui/styles.py — QSS Stylesheet
MAIN_STYLE = """
QMainWindow {
    background-color: #f8f9fa;
}
QGroupBox {
    font-weight: bold;
    border: 1px solid #dee2e6;
    border-radius: 8px;
    margin-top: 12px;
    padding-top: 16px;
}
QPushButton {
    background-color: #4361ee;
    color: white;
    border: none;
    border-radius: 6px;
    padding: 8px 20px;
    font-size: 14px;
    font-weight: 500;
}
QPushButton:hover {
    background-color: #3a56d4;
}
QPushButton:pressed {
    background-color: #2f48b8;
}
QPushButton:disabled {
    background-color: #adb5bd;
}
QPushButton#exportBtn {
    background-color: #2d6a4f;
}
QPushButton#exportBtn:hover {
    background-color: #245a42;
}
QTableWidget {
    border: 1px solid #dee2e6;
    border-radius: 4px;
    gridline-color: #e9ecef;
    selection-background-color: #e7f1ff;
    font-size: 12px;
}
QTableWidget::item {
    padding: 4px 8px;
}
QHeaderView::section {
    background-color: #495057;
    color: white;
    padding: 6px;
    border: none;
    font-weight: bold;
}
QProgressBar {
    border: 1px solid #dee2e6;
    border-radius: 4px;
    text-align: center;
    height: 24px;
}
QProgressBar::chunk {
    background-color: #4361ee;
    border-radius: 3px;
}
QTextEdit#logPanel {
    background-color: #212529;
    color: #a8e6cf;
    font-family: 'Consolas', 'Courier New', monospace;
    font-size: 11px;
    border-radius: 4px;
    padding: 8px;
}
"""
```

### 3.6 Threading (QThread)

ทุก heavy computation ต้องรันใน QThread เพื่อไม่ block GUI:

```python
class CalculationWorker(QThread):
    """Worker thread for running the calculation pipeline."""
    progress = pyqtSignal(int, str)     # (percent, message)
    result_ready = pyqtSignal(object)   # list[IndexResult]
    error = pyqtSignal(str)             # error message
    log_message = pyqtSignal(str)       # log line
    
    def __init__(self, df, dimensions, subgroups):
        super().__init__()
        self.df = df
        self.dimensions = dimensions
        self.subgroups = subgroups
    
    def run(self):
        try:
            pipeline = IndexCalculationPipeline(self.df, self.dimensions)
            results = []
            total = len(self.subgroups)
            
            for i, subgroup in enumerate(self.subgroups):
                self.log_message.emit(f"Processing: {subgroup.name}...")
                self.progress.emit(int((i / total) * 100), subgroup.name)
                
                mask = pipeline.build_mask(subgroup)
                n = mask.sum()
                if n < 30:
                    self.log_message.emit(f"  Skipped (n={n} < 30)")
                    continue
                
                result = pipeline.run_for_subgroup(mask)
                results.append(result)
                self.log_message.emit(
                    f"  Done: Index={result.index_value:.3f}, n={result.n}"
                )
            
            self.progress.emit(100, "Complete!")
            self.result_ready.emit(results)
        except Exception as e:
            self.error.emit(str(e))
```

---

## 4. Statistical Implementation Notes

### 4.1 Factor Analysis Details

เนื่องจาก Python ไม่มี EQUAMAX rotation ใน `factor_analyzer` โดยตรง (มีแค่ varimax, promax, oblimin, quartimax) วิธีแก้:

**Option A**: ใช้ `varimax` แทน — ผลลัพธ์ใกล้เคียงเพราะข้อมูล 6 variables
**Option B**: Implement equamax เอง — EQUAMAX = gamma ที่อยู่ระหว่าง varimax (gamma=1) กับ quartimax (gamma=0), equamax ใช้ gamma = p/2 โดย p = จำนวน factors

```python
# Equamax implementation hint:
# Use scipy.optimize to minimize the equamax criterion
# gamma = n_factors / 2 (สำหรับ equamax)
# Can also use R's `GPArotation` via rpy2 if needed
```

**Option C (แนะนำ)**: เนื่องจากเรา extract 6 factors จาก 6 variables → factor scores จะเป็น orthogonal อยู่แล้ว (เพราะ PCA ให้ orthogonal components) ดังนั้น rotation ไม่เปลี่ยน factor scores ถ้าเราใช้ save AR (Anderson-Rubin scores) ซึ่งเป็น orthogonal เสมอ

### 4.2 Standardized Beta

SPSS ให้ standardized coefficients โดยอัตโนมัติ ใน Python ต้องคำนวณเอง:

```python
def standardized_betas(X: np.ndarray, y: np.ndarray, betas: np.ndarray) -> np.ndarray:
    """Compute standardized regression coefficients."""
    sx = np.std(X, axis=0, ddof=1)
    sy = np.std(y, ddof=1)
    return betas * (sx / sy)
```

หรือใช้ `sklearn.preprocessing.StandardScaler` ก่อน fit:
```python
from sklearn.preprocessing import StandardScaler
scaler = StandardScaler()
X_std = scaler.fit_transform(X)
y_std = scaler.fit_transform(y.values.reshape(-1, 1)).ravel()
model = sm.OLS(y_std, sm.add_constant(X_std)).fit()
std_betas = model.params[1:]  # These ARE standardized betas
```

### 4.3 Pearson Correlation for Output

```python
from scipy.stats import pearsonr

def compute_correlations(index_series: pd.Series, items_df: pd.DataFrame) -> pd.DataFrame:
    """Compute Pearson r, p-value, and N for each item vs Index."""
    results = []
    for col in items_df.columns:
        valid = index_series.notna() & items_df[col].notna()
        n = valid.sum()
        if n > 2:
            r, p = pearsonr(index_series[valid], items_df[col][valid])
        else:
            r, p = np.nan, np.nan
        results.append({'item': col, 'r': r, 'p': p, 'n': n})
    return pd.DataFrame(results)
```

### 4.4 Missing Value Handling

- ใช้ **listwise deletion** ต่อ subgroup (consistent กับ SPSS FACTOR /MISSING LISTWISE)
- สำหรับ Dimension means ใช้ `np.nanmean` (คำนวณจาก valid items เท่านั้น)
- สำหรับ correlation: ใช้ pairwise deletion (คำนวณ N แยกต่อ item pair)

---

## 5. Excel Export Format

### 5.1 Sheet Structure

```python
# export_service.py

def export_to_excel(results: list[IndexResult], correlations_q2, correlations_q3, output_path: Path):
    with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
        # Sheet 1: Correlations Q2
        write_correlation_sheet(writer, 'Sheet1', correlations_q2, subgroup_columns_q2)
        
        # Sheet 2: Correlations Q3  
        write_correlation_sheet(writer, 'Sheet2', correlations_q3, subgroup_columns_q3)
        
        # Sheet 3: Index Summary
        write_index_sheet(writer, 'Index', results)
```

### 5.2 Correlation Sheet Layout

```
Row 1: Headers → "Correlations" | "" | "Total" | "Japan" | "China" | ...
Row 2: Sub-header → "" | "" | "Index" | "Index" | "Index" | ...
Row 3-5: Index self-correlation (r=1, sig=blank, N)
Row 6-8: PWQ2#1 (r, sig, N)
Row 9-11: PWQ2#2 (r, sig, N)
...

Column A: Item label (merged 3 rows)
Column B: Statistic type ("Pearson Correlation" / "Sig. (2-tailed)" / "N")
Column C onwards: Subgroup values
```

### 5.3 Index Sheet Layout

```
Row 1: "" | "" | "score (1-7)" merged C-H | "Weight" merged I-N
Row 2: "" | "Index" | "Physical" | "Mental" | "Social" | "Financial" | "Environmental" | "Purpose & spiritual" | (same 6 for weights)
Row 3: "Total" | 5.133 | 5.19 | 5.13 | 5.08 | 5.12 | 5.15 | 5.15 | 0.143 | 0.199 | 0.182 | 0.197 | 0.110 | 0.170
Row 4: "Japan" | 4.426 | ...
...
```

---

## 6. Edge Cases & Error Handling

1. **ไฟล์ SPSS ไม่มีตัวแปรที่ต้องการ** → แสดง error ชัดเจน บอกว่าขาดตัวแปรอะไร
2. **Subgroup มี N < 30** → ข้าม subgroup นั้น แสดง warning ใน log
3. **Factor Analysis fail** (e.g., singular matrix) → fallback ใช้ PCA ไม่ rotate
4. **ตัวแปร Q3 ไม่มีในไฟล์** → ข้าม Sheet 2 ของ output
5. **Missing values มากเกินไป** (>50% ของ subgroup) → warn user
6. **ค่า Q1 ไม่อยู่ในช่วง 1-7** → validate ก่อน recode
7. **Regression: multicollinearity สูง** (VIF > 10) → ไม่เกิดเพราะใช้ factor scores (orthogonal) แต่ยังต้อง check

---

## 7. Dependencies (pyproject.toml)

```toml
[project]
name = "wellbeing-index-calculator"
version = "1.0.0"
requires-python = ">=3.11"
dependencies = [
    "PyQt6>=6.6",
    "pyreadstat>=1.2",
    "pandas>=2.1",
    "numpy>=1.26",
    "factor-analyzer>=0.5",
    "statsmodels>=0.14",
    "scipy>=1.12",
    "openpyxl>=3.1",
]

[project.optional-dependencies]
dev = [
    "pytest>=8.0",
    "ruff>=0.3",
    "pyinstaller>=6.0",
]

[project.scripts]
wellbeing = "wellbeing.main:main"
```

---

## 8. Build & Distribution

สำหรับแจกจ่ายให้ทีม (Windows):

```bash
# Build standalone .exe
pyinstaller --onefile --windowed \
    --name "WellbeingIndexCalculator" \
    --icon assets/icon.ico \
    --add-data "assets;assets" \
    src/wellbeing/main.py
```

---

## 9. Testing Checklist

- [ ] Load .sav file → verify variable detection
- [ ] Recode Q1 → verify NQ1 = 8 - Q1
- [ ] Dim means → compare with SPSS COMPUTE MEAN output
- [ ] Factor Analysis → compare factor scores with SPSS output (tolerance ±0.01)
- [ ] Regression → compare R², Beta with SPSS output (tolerance ±0.001)
- [ ] Weights → verify sum = 1.0
- [ ] Index → compare with SPSS COMPUTE Index output
- [ ] Subgroup results → spot-check 3-4 subgroups against SPSS
- [ ] Excel output → verify layout matches reference Output.xlsx
- [ ] Edge case: file with no Q3 → Sheet2 should be empty/skipped
- [ ] Edge case: subgroup N < 30 → should be skipped with warning
- [ ] GUI responsiveness during calculation → UI must not freeze
