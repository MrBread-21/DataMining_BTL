# Bàn giao dự án: Phân khúc khách hàng, dự đoán rời bỏ và gợi ý sản phẩm

Đồ án môn **Khai phá dữ liệu**. Tài liệu này giúp người nhận bàn giao nắm được dự án làm gì, mỗi bước được làm như thế nào và chạy lại code ra sao.

- **Trạng thái:** cả 4 notebook đã hoàn chỉnh và chạy được từ đầu đến cuối. Output, 21 biểu đồ và 6 file dữ liệu trung gian đã có sẵn.
- **Còn lại:** viết báo cáo và làm slide (xem [mục 10](#10-việc-còn-lại-báo-cáo-và-slide)).

---

## 1. Tổng quan

Dự án phân tích dữ liệu giao dịch của một cửa hàng quà tặng trực tuyến ở Anh, nhằm:

1. **Phân khúc khách hàng** theo mô hình RFM bằng phân cụm K-Means.
2. **Dự đoán khách hàng sắp rời bỏ** (churn) trong 3 tháng tới bằng các mô hình phân lớp.
3. **Tìm luật kết hợp theo từng phân khúc** để gợi ý sản phẩm cho khách có nguy cơ rời bỏ.
4. Tổng hợp thành **đề xuất chiến lược kinh doanh** cho từng nhóm khách.

Các bước nối với nhau thành một mạch:

- Nhãn cụm của bước 1 là đặc trưng đầu vào của mô hình churn ở bước 2.
- Xác suất churn của bước 2 dùng để chọn khách cần giữ chân.
- Luật kết hợp của cụm mà khách thuộc về cho biết nên gợi ý cho khách đó sản phẩm gì.

### Kết quả chính

| Bước | Kết quả |
|---|---|
| Tiền xử lý | Từ 1.067.371 dòng gốc còn 770.688 dòng sạch. Tập nghiên cứu gồm 4.299 khách, tỷ lệ churn 49,2% (hai lớp gần cân bằng) |
| Phân cụm | K-Means, k = 4, ra 4 cụm: **VIP, Khách mới, Có nguy cơ, Ngủ đông**. VIP chiếm 18% số khách nhưng tạo ra 69% doanh thu. Tỷ lệ churn tăng dần từ 14% (VIP) đến 71% (Ngủ đông) |
| Dự đoán churn | XGBoost được chọn theo cross-validation: ROC-AUC 0,77 trên CV, 0,73 trên tập test. Bốn mô hình chênh nhau không đáng kể. **Recency là yếu tố quan trọng nhất**; khoảng 4 tháng không mua là ngưỡng nên cảnh báo |
| Luật kết hợp và gợi ý | 1.105 luật theo cụm và 1.387 luật mua kèm toàn cửa hàng, phần lớn có dạng "mua đủ bộ, đủ màu". Có gợi ý cho 1.076 khách nguy cơ cao, mỗi gợi ý kèm lý do bằng lời; gần 80% khách có gợi ý dựa trên sản phẩm đã mua. Gợi ý theo luật trúng 6,5% (luật của cụm) và 5,9% (luật toàn cửa hàng), so với 3,5% của hàng bán chạy |

---

## 2. Dữ liệu

- **Nguồn:** bộ Online Retail II, UCI Machine Learning Repository (dataset id 502).
- **Phạm vi:** giao dịch từ 01/12/2009 đến 09/12/2011, gồm 1.067.371 dòng và 8 cột.
- **File đang dùng:** `data/raw/online_retail_II.csv` (khoảng 91 MB). File này là hai sheet `Year 2009-2010` và `Year 2010-2011` của file Excel gốc nối lại với nhau. **Không chỉnh sửa file này.**

| Cột | Ý nghĩa |
|---|---|
| `Invoice` | Mã hóa đơn. Mã bắt đầu bằng `C` là hóa đơn hủy, bằng `A` là bút toán điều chỉnh |
| `StockCode` | Mã sản phẩm |
| `Description` | Tên sản phẩm |
| `Quantity` | Số lượng |
| `InvoiceDate` | Thời điểm giao dịch |
| `Price` | Đơn giá (GBP) |
| `Customer ID` | Mã khách hàng (khoảng 23% số dòng bị thiếu) |
| `Country` | Quốc gia của khách |

**Nếu chỉ có file Excel gốc** (`online_retail_II.xlsx`, tải từ UCI): đặt file vào `data/raw/`, rồi thay ô đọc dữ liệu ở mục 1 của notebook 01 bằng đoạn sau. Đoạn này cần `openpyxl` và đọc khá lâu (vài phút). Lưu ý: đoạn này **chưa được chạy thử**, vì máy hiện chỉ có file CSV.

```python
sheets = pd.read_excel("../data/raw/online_retail_II.xlsx", sheet_name=None,
                       dtype={"Invoice": str, "StockCode": str})
df_raw = pd.concat(sheets.values(), ignore_index=True)
```

---

## 3. Cấu trúc thư mục

```
DataMining/
|-- BAN_GIAO_DU_AN.md            # tài liệu này
|-- requirements.txt             # danh sách thư viện
|-- data/
|   |-- raw/
|   |   `-- online_retail_II.csv # dữ liệu gốc (91 MB), không chỉnh sửa
|   `-- processed/               # dữ liệu trung gian do các notebook sinh ra (73 MB)
|-- notebooks/
|   |-- 01_tien_xu_ly_va_eda.ipynb
|   |-- 02_phan_cum_rfm.ipynb
|   |-- 03_du_doan_churn.ipynb
|   `-- 04_luat_ket_hop.ipynb
`-- figures/                     # 21 biểu đồ PNG để chèn vào báo cáo
```

Mỗi notebook chỉ đọc dữ liệu do notebook trước xuất ra trong `data/processed/` (riêng notebook 01 đọc file gốc). Vì vậy không notebook nào phải chạy lại bước của notebook khác.

| Notebook | Đọc | Ghi ra `data/processed/` |
|---|---|---|
| 01 Tiền xử lý và EDA | `raw/online_retail_II.csv` | `transactions_clean.csv`, `customers.csv` |
| 02 Phân cụm RFM | `transactions_clean.csv`, `customers.csv` | `customers_clustered.csv` |
| 03 Dự đoán churn | `transactions_clean.csv`, `customers_clustered.csv` | `churn_predictions.csv` |
| 04 Luật kết hợp | `transactions_clean.csv`, `customers_clustered.csv`, `churn_predictions.csv` | `rules_by_cluster.csv`, `recommendations.csv` |

Thư mục `data/` khá nặng (khoảng 164 MB). Nếu nộp bài hoặc đưa lên GitHub, nên bỏ thư mục này ra; người nhận tự đặt file gốc vào `data/raw/` rồi chạy lại.

---

## 4. Cài đặt môi trường

Cần **Python 3.10 trở lên**. Dự án đã chạy với Python 3.10.11 trên Windows 11.

Mở terminal tại thư mục gốc của dự án (PowerShell):

```powershell
python -m venv .venv
.venv\Scripts\Activate.ps1
pip install -r requirements.txt
```

Các phiên bản đã chạy thử:

| Thư viện | Phiên bản | Ghi chú |
|---|---|---|
| pandas | 2.3.3 | |
| numpy | 1.26.4 | Chưa chạy thử với numpy 2 |
| scikit-learn | 1.7.2 | |
| scipy | 1.15.3 | Ma trận thưa để đếm cặp sản phẩm ở notebook 04 (cài kèm scikit-learn) |
| xgboost | 3.2.0 | |
| mlxtend | 0.23.4 | Bắt buộc từ 0.23.4, vì `association_rules()` cần tham số `num_itemsets` |
| shap | 0.49.1 | Chỉ dùng để vẽ biểu đồ SHAP (xem [mục 9](#9-vấn-đề-đã-biết)) |
| matplotlib | 3.10.8 | |
| nbconvert | 7.17.0 | Cài kèm gói `jupyter` |

`seaborn` và `imbalanced-learn` có trong `requirements.txt` nhưng hiện chưa notebook nào dùng. `openpyxl` chỉ cần khi đọc file Excel gốc.

---

## 5. Cách chạy

**Luôn chạy theo thứ tự 01 -> 02 -> 03 -> 04**, vì notebook sau đọc file do notebook trước tạo ra.

### Cách 1: mở bằng Jupyter

```powershell
cd notebooks
python -m notebook
```

Mở lần lượt từng notebook và chọn **Kernel -> Restart Kernel and Run All Cells**.

### Cách 2: chạy bằng dòng lệnh

Chạy từ thư mục `notebooks/`, vì đường dẫn trong code là đường dẫn tương đối (ví dụ `../data/processed/`). Trên Windows, lệnh `jupyter` thường không có trong PATH, nên dùng `python -m nbconvert`:

```powershell
cd notebooks
python -m nbconvert --to notebook --execute --inplace 01_tien_xu_ly_va_eda.ipynb
python -m nbconvert --to notebook --execute --inplace 02_phan_cum_rfm.ipynb
python -m nbconvert --to notebook --execute --inplace 03_du_doan_churn.ipynb
python -m nbconvert --to notebook --execute --inplace 04_luat_ket_hop.ipynb
```

`--inplace` ghi kết quả chạy đè lên chính notebook đó.

### Thời gian chạy (đo trên máy hiện tại)

| Notebook | Thời gian |
|---|---|
| 01 Tiền xử lý và EDA | khoảng 30 giây |
| 02 Phân cụm RFM | khoảng 30 giây |
| 03 Dự đoán churn | khoảng 50 giây |
| 04 Luật kết hợp | khoảng 1 phút |
| **Tổng** | **khoảng 3 phút** |

Khi chạy bằng dòng lệnh trên Windows, terminal có thể in `RuntimeWarning: Proactor event loop does not implement add_reader...`. Cảnh báo này đến từ thư viện giao tiếp của Jupyter, không ảnh hưởng kết quả và không xuất hiện trong notebook.

### Kết quả có lặp lại được không?

Có. Mọi chỗ có yếu tố ngẫu nhiên đều cố định `RANDOM_STATE = 42`. Ở notebook 04, các luật và sản phẩm bằng điểm nhau được xếp tiếp theo tiêu chí phụ (lift -> confidence -> tên sản phẩm). Đã kiểm tra: chạy lại nhiều lần, kể cả với các hash seed khác nhau của Python, đều cho đúng cùng số liệu.

---

## 6. Cách làm

### 6.1. Định nghĩa rời bỏ (churn) và các mốc thời gian

Đây là phần cốt lõi của dự án. Thay đổi bất kỳ mốc nào ở đây đều làm thay đổi toàn bộ kết quả phía sau.

| Khái niệm | Giá trị |
|---|---|
| Giai đoạn quan sát | Từ đầu dữ liệu (01/12/2009) đến hết 09/09/2011. **Mọi đặc trưng chỉ tính từ giai đoạn này** |
| Giai đoạn dự đoán | Từ 10/09/2011 đến 09/12/2011 (3 tháng cuối) |
| Ngày tham chiếu (`REF_DATE`) | 10/09/2011, dùng để tính Recency và Tenure |
| Tập khách nghiên cứu | Khách có ít nhất 1 giao dịch trong 12 tháng cuối của giai đoạn quan sát (từ 10/09/2010): **4.299 khách** |
| Nhãn | `churn = 1` nếu khách **không có giao dịch nào** trong giai đoạn dự đoán, ngược lại `churn = 0` |

- **Vì sao giới hạn ở khách mua trong 12 tháng cuối:** 949 khách đã ngưng mua hơn một năm bị loại, vì những khách này coi như đã mất (89,3% trong số họ không quay lại), không cần dự đoán.
- **Chống rò rỉ dữ liệu (data leakage):** không dùng dữ liệu của giai đoạn dự đoán để tính đặc trưng, kể cả khi làm sạch dữ liệu. Ngoại lệ duy nhất là mục 5.5 của notebook 04: dữ liệu giai đoạn dự đoán được dùng để **đánh giá** gợi ý, không dùng để xây dựng gợi ý.

### 6.2. Notebook 01: Tiền xử lý và EDA

**Các bước làm sạch** (notebook in số dòng còn lại sau mỗi bước):

| Bước | Nội dung | Số dòng còn lại |
|---|---|---|
| 1 | Bỏ dòng trùng lặp (34.335 dòng, trong đó 22.844 dòng do hai sheet chồng lấn nhau từ 01/12 đến 09/12/2010) | 1.033.036 |
| 2 | Bỏ dòng thiếu `Customer ID`, đổi tên cột thành `CustomerID`, ép về kiểu số nguyên | 797.885 |
| 3a | Bỏ đơn mua gốc đã bị hủy (xem giải thích bên dưới) | 791.852 |
| 3b | Bỏ hóa đơn hủy (`C`) và hóa đơn điều chỉnh (`A`) | 773.462 |
| 4 | Bỏ dòng có `Quantity <= 0` hoặc `Price <= 0` | 773.393 |
| 5 | Chỉ giữ mã sản phẩm thật: `StockCode` bắt đầu bằng 5 chữ số (regex `^\d{5}`), loại phí ship, giảm giá, điều chỉnh thủ công... | 770.688 |
| 6 | Chuẩn hóa tên sản phẩm: bỏ khoảng trắng thừa, mỗi `StockCode` dùng tên xuất hiện nhiều nhất | 770.688 |
| 7 | Tạo cột `TotalPrice = Quantity * Price` | 770.688 |

**Bước 3a** không có trong cách làm sạch thông thường nhưng cần thiết:

- **Vấn đề:** nếu chỉ bỏ dòng hủy thì đơn mua gốc vẫn còn, làm doanh thu bị thổi phồng. Ví dụ: một đơn 80.995 sản phẩm (£168K) bị hủy ngay trong ngày.
- **Cách làm:** với mỗi dòng hủy, tìm dòng mua gần nhất trước đó của cùng khách, cùng sản phẩm, cùng số lượng (dùng `pd.merge_asof`), rồi loại dòng mua đó.
- **Điều kiện:** chỉ ghép khi đơn hủy và đơn gốc nằm **cùng một giai đoạn** so với `REF_DATE`. Nếu không, ta đã dùng thông tin tương lai để sửa dữ liệu quá khứ. 331 dòng hủy ở giai đoạn dự đoán có đơn gốc ở giai đoạn quan sát nên không được ghép.
- **Kết quả:** loại 6.033 dòng mua, khoảng £666K doanh thu ảo.

**Bước 6** giúp luật kết hợp ở notebook 04 không bị tách đôi chỉ vì cùng một sản phẩm có nhiều cách ghi tên. Sau bước này còn 4.611 mã sản phẩm với 4.574 tên.

**EDA:** doanh thu theo tháng (tính mùa vụ rõ, đỉnh vào tháng 11), top 10 sản phẩm, top 10 quốc gia (Anh chiếm 83,5% doanh thu), phân phối số hóa đơn mỗi khách, phân phối giá trị đơn hàng (lệch phải mạnh, trung vị £301).

### 6.3. Notebook 02: Phân cụm RFM

1. **Tính RFM** cho mỗi khách từ giao dịch giai đoạn quan sát:
   - `Recency`: số ngày từ lần mua cuối đến 10/09/2011.
   - `Frequency`: số hóa đơn khác nhau.
   - `Monetary`: tổng `TotalPrice`.
2. **Biến đổi:** `log1p` cho cả 3 biến (F và M lệch phải rất mạnh), sau đó chuẩn hóa bằng `StandardScaler`.
3. **Chọn k:** thử k = 2..10 với Elbow, Silhouette và Davies-Bouldin, chọn **k = 4** (Silhouette 0,361, cao nhất trong các k ≥ 3; Davies-Bouldin 0,907). k = 2 có Silhouette cao hơn nhưng quá thô để xây dựng chiến lược.
4. **So sánh thuật toán:** K-Means tốt hơn GMM và Agglomerative (Ward) trên cả Silhouette và Davies-Bouldin, nên được chọn.
5. **Đặt tên cụm** bằng quy tắc viết thành code, để kết quả không phụ thuộc vào thứ tự nhãn ngẫu nhiên của K-Means:
   - Cụm có Frequency trung bình cao nhất là **VIP**.
   - Trong các cụm còn lại, cụm có Recency thấp nhất là **Khách mới**.
   - Trong hai cụm còn lại, cụm có Frequency cao hơn là **Có nguy cơ**, cụm kia là **Ngủ đông**.
   - Cột `Cluster` được đánh số lại: 0 = VIP, 1 = Khách mới, 2 = Có nguy cơ, 3 = Ngủ đông.

| Cụm | Số khách | % doanh thu | Tỷ lệ churn | Recency TB (ngày) | Frequency TB | Monetary TB |
|---|---|---|---|---|---|---|
| VIP | 790 | 68,8% | 14,1% | 28 | 20,3 | £11.089 |
| Khách mới | 563 | 3,6% | 42,3% | 24 | 2,9 | £815 |
| Có nguy cơ | 1.410 | 22,8% | 47,5% | 161 | 5,6 | £2.060 |
| Ngủ đông | 1.536 | 4,8% | 71,4% | 232 | 1,6 | £396 |

Notebook có thêm một phép thử: giữ Recency ở thang gốc thay vì `log1p`. Phương án này tách tỷ lệ churn của hai cụm giữa rõ hơn nhưng không tạo ra được cụm "Khách mới", nên phương án chính vẫn là `log1p` cho cả 3 biến.

### 6.4. Notebook 03: Dự đoán churn

**Đặc trưng** (chỉ từ giai đoạn quan sát):

- `Recency`, `Frequency`, `Monetary`.
- `Tenure`: số ngày từ lần mua đầu tiên đến 10/09/2011.
- `AvgOrderValue`: giá trị trung bình mỗi hóa đơn.
- `NumProducts`: số sản phẩm khác nhau đã mua.
- `IsUK`: 1 nếu khách ở Anh, 0 nếu ở nước khác.
- `Cluster_*`: nhãn cụm từ notebook 02, mã hóa one-hot.

F, M, AvgOrderValue và NumProducts được biến đổi `log1p`.

**Huấn luyện:**

- Chia train/test theo tỷ lệ 80/20 với `stratify=churn` (3.439 / 860 khách).
- Bốn mô hình: Logistic Regression (mốc so sánh), Decision Tree, Random Forest, XGBoost.
- Tinh chỉnh bằng `GridSearchCV` nhỏ, 5-fold `StratifiedKFold`, tiêu chí ROC-AUC.
- Với Logistic Regression, scaler nằm trong `Pipeline`, nên chỉ được fit trên dữ liệu train.
- **Mô hình được chọn theo điểm CV, không theo điểm test.**
- Không dùng SMOTE hay `class_weight`, vì hai lớp đã gần cân bằng (49,2% churn).

| Mô hình | ROC-AUC CV | ROC-AUC test | F1 test |
|---|---|---|---|
| Logistic Regression | 0,768 | 0,729 | 0,667 |
| Decision Tree | 0,757 | 0,720 | 0,688 |
| Random Forest | 0,769 | 0,732 | 0,667 |
| **XGBoost** (được chọn: `max_depth=2`, `n_estimators=300`, `learning_rate=0.03`) | **0,771** | 0,730 | 0,667 |

**Giải thích mô hình:**

- **Permutation importance:** Recency quan trọng nhất, sau đó là Frequency và Tenure. Nhãn cụm và quốc gia gần như không đóng góp.
- **SHAP:** cho biết chiều tác động của từng đặc trưng.
- **Hệ số Logistic Regression.**
- **Cây quyết định:** huấn luyện lại trên giá trị gốc (không log) để ngưỡng dễ đọc. Quy tắc nổi bật: khách mua không quá 6 lần và đã hơn khoảng 4 tháng (126 ngày) không mua có tỷ lệ rời bỏ 71%.

**Dự đoán cho toàn bộ khách:**

- Khách thuộc tập train nhận xác suất *out-of-fold* (dùng `cross_val_predict`).
- Khách thuộc tập test nhận dự đoán thông thường.
- Như vậy không xác suất nào được tính trên chính dữ liệu mà mô hình đã học.

**Lưu ý quan trọng:** ở ngưỡng 0,5, mô hình chỉ cảnh báo được 1 khách VIP, dù thực tế có khoảng 110 khách VIP rời bỏ (xác suất của khách VIP đều thấp). Tuy vậy, bên trong cụm VIP mô hình vẫn **xếp hạng** khách khá tốt (ROC-AUC 0,75). Đây là lý do notebook 04 chọn khách theo thứ hạng trong từng cụm.

### 6.5. Notebook 04: Luật kết hợp và gợi ý sản phẩm

**Khai phá luật theo từng cụm:**

1. Lấy giao dịch giai đoạn quan sát và gắn nhãn cụm của khách.
2. Tạo ma trận giỏ hàng: mỗi dòng là một hóa đơn, mỗi cột là một sản phẩm (`Description`).
3. Chọn `min_support = max(1%, 20 / số hóa đơn của cụm)`, để cụm nhỏ vẫn cần ít nhất 20 hóa đơn cho một luật.
4. Bỏ trước các sản phẩm có support thấp hơn `min_support` (theo tính chất Apriori), nhưng vẫn giữ đủ mọi hóa đơn để support được tính đúng mẫu số.
5. Chạy **FP-Growth** (`mlxtend`), giữ các luật có `lift > 1` và `confidence >= 0,5`.
6. Sắp xếp luật theo lift, rồi confidence, rồi tên sản phẩm.

Kết quả: VIP 560 luật, Khách mới 92, Có nguy cơ 339, Ngủ đông 114. Các cụm được so sánh bằng hệ số Jaccard giữa hai tập luật (từ 0,07 đến 0,32).

**Chọn khách nguy cơ cao:** 25% khách có xác suất churn cao nhất **trong từng cụm**, tổng cộng 1.076 khách. Nhóm được chọn có tỷ lệ churn thực tế cao hơn rõ rệt so với mức chung của cụm (ví dụ VIP: 28,3% so với 14,1%).

**Luật mua kèm toàn cửa hàng** (dùng khi luật của cụm không đủ):

- **Vấn đề:** luật của cụm cần ít nhất 1% số hóa đơn của cụm, nên nhiều sản phẩm ít phổ biến không có luật nào. Nếu chỉ dùng luật của cụm, 505 trên 1.076 khách nguy cơ cao (47%) không nhận được gợi ý nào dựa trên sản phẩm đã mua.
- **Cách làm:** tìm thêm luật một sản phẩm -> một sản phẩm trên toàn bộ hóa đơn giai đoạn quan sát (gộp mọi cụm), với cùng tiêu chuẩn: ít nhất 20 hóa đơn chứa cả hai sản phẩm, confidence ≥ 0,5, lift > 1.
- **Vì sao không dùng FP-Growth ở đây:** với ngưỡng thấp như vậy, FP-Growth chạy hơn 10 phút chưa xong. Notebook đếm trực tiếp số hóa đơn chứa từng cặp sản phẩm bằng phép nhân ma trận thưa (`basket.T @ basket`, dùng `scipy.sparse`). Cách này cho đúng các luật 2 sản phẩm mà FP-Growth tìm được, chỉ mất vài giây.
- **Kết quả:** 1.387 luật, vế trái gồm 839 sản phẩm (luật theo cụm chỉ có 186 sản phẩm ở vế trái).

**Thuật toán gợi ý** (5 sản phẩm mỗi khách), điền lần lượt từ 3 nguồn:

1. **Luật của cụm** mà khách thuộc về, theo thứ tự lift giảm dần. Nếu mọi sản phẩm ở vế trái của luật khách đều đã mua, gợi ý các sản phẩm ở vế phải mà khách chưa mua.
2. Nếu chưa đủ 5 sản phẩm, dùng tiếp **luật toàn cửa hàng** theo cùng cách.
3. Nếu vẫn chưa đủ, bổ sung bằng **hàng bán chạy nhất trong cụm** mà khách chưa mua. Đây là gợi ý theo độ phổ biến, không dựa trên sản phẩm khách đã mua.

Các cột của bảng gợi ý:

- Cột `Source` ghi nguồn: `Luật của cụm`, `Luật toàn cửa hàng` hoặc `Bán chạy trong cụm`.
- Cột `BasedOn` ghi sản phẩm khách đã mua làm căn cứ.
- Cột `Reason` giải thích lý do của từng gợi ý bằng lời, do các hàm `explain_rule()` và `fallback_reason()` tạo ra:
  - **Gợi ý từ luật:** nêu sản phẩm khách đã mua, luật đó là của cụm hay của toàn cửa hàng, tỷ lệ hóa đơn có các sản phẩm đó cũng có sản phẩm được gợi ý (confidence), mức tăng so với tỷ lệ chung (lift) và số hóa đơn làm căn cứ.
  - **Gợi ý từ hàng bán chạy:** nêu cụ thể khách đã mua gì. Code xác định đúng một trong ba trường hợp: không luật nào khớp với sản phẩm đã mua; có luật khớp nhưng luật chỉ gợi ý sản phẩm khách đã có; hoặc luật chỉ cho được vài gợi ý. Sau đó nêu thứ hạng và độ phổ biến của sản phẩm trong cụm.

**Kết quả gợi ý:** gần 80% khách nguy cơ cao có ít nhất 1 gợi ý dựa trên sản phẩm đã mua. 218 khách (20%) chỉ nhận hàng bán chạy; phần lớn họ thuộc cụm Ngủ đông và có trung vị chỉ 4 sản phẩm đã mua.

**Đánh giá gợi ý:**

- So sánh gợi ý với những gì khách thực sự mua trong giai đoạn dự đoán (chỉ để đánh giá).
- Tỷ lệ trúng: luật của cụm 6,5%, luật toàn cửa hàng 5,9%, hàng bán chạy 3,5%.
- Luật của cụm trúng cao hơn luật toàn cửa hàng ở VIP, Khách mới và Có nguy cơ. Ngoại lệ là cụm Ngủ đông: ở đây hàng bán chạy trúng nhiều nhất (5,7%).

Cuối notebook 04 có phần **đề xuất chiến lược cho từng cụm** và **kết luận chung** của dự án (kết quả, hạn chế, hướng phát triển). Mỗi cụm được trình bày theo 4 phần: tình hình, hành động đề xuất, căn cứ (số liệu từ các notebook giải thích vì sao đề xuất như vậy) và sản phẩm gợi ý.

Mọi bảng và biểu đồ trong notebook 04 đều có phần **Cách đọc** giải thích ý nghĩa các cột, trước phần **Nhận xét**.

---

## 7. Mô tả các file đầu ra

Các file có tiếng Việt được ghi với encoding `utf-8-sig`, nên mở bằng Excel không bị lỗi font.

| File | Kích thước | Các cột |
|---|---|---|
| `transactions_clean.csv` | 770.688 × 9 | Giao dịch đã làm sạch: `Invoice`, `StockCode`, `Description`, `Quantity`, `InvoiceDate`, `Price`, `CustomerID`, `Country`, `TotalPrice` (toàn bộ thời gian, gồm cả giai đoạn dự đoán) |
| `customers.csv` | 4.299 × 2 | `CustomerID`, `churn` |
| `customers_clustered.csv` | 4.299 × 7 | `CustomerID`, `Recency`, `Frequency`, `Monetary`, `Cluster` (0 đến 3), `ClusterName`, `churn` |
| `churn_predictions.csv` | 4.299 × 6 | `CustomerID`, `ClusterName`, `churn_proba` (xác suất rời bỏ), `churn_pred` (nhãn dự đoán ở ngưỡng 0,5), `churn` (nhãn thật), `Set` (`train` hoặc `test`) |
| `rules_by_cluster.csv` | 1.105 × 8 | `ClusterName`, `Antecedents`, `Consequents` (nhiều sản phẩm nối bằng ` + `), `AntecedentSupport`, `ConsequentSupport`, `Support`, `Confidence`, `Lift` |
| `recommendations.csv` | 5.380 × 10 | `CustomerID`, `ClusterName`, `churn_proba`, `Rank` (1 đến 5), `RecommendedProduct`, `Source` (`Luật của cụm`, `Luật toàn cửa hàng` hoặc `Bán chạy trong cụm`), `BasedOn`, `Confidence`, `Lift` (ba cột này để trống nếu là hàng bán chạy), `Reason` (lý do gợi ý viết bằng lời) |

Mỗi biểu đồ trong `figures/` có tiền tố là số của notebook sinh ra nó (ví dụ `01_doanh_thu_theo_thang.png`, `03_duong_roc.png`).

---

## 8. Quy ước và lưu ý khi chỉnh sửa

### Quy ước

- Đường dẫn trong code là đường dẫn tương đối, tính từ thư mục `notebooks/`.
- `RANDOM_STATE = 42` ở mọi chỗ có yếu tố ngẫu nhiên.
- Ô markdown viết bằng tiếng Việt, gồm mục đích và nhận xét kết quả của từng bước, có thể chép thẳng vào báo cáo. Tên biến, tên hàm và chú thích code viết bằng tiếng Anh.
- Biểu đồ được lưu bằng hàm `save_fig()` khai báo ở đầu mỗi notebook.
- Màu của mỗi cụm cố định trong mọi biểu đồ: VIP xanh dương, Khách mới cam, Có nguy cơ xanh lá, Ngủ đông vàng. Nếu thêm biểu đồ mới, nên giữ đúng màu này.

### Lưu ý khi chỉnh sửa

1. **Các con số trong ô nhận xét được viết tay** theo kết quả hiện tại, kể cả phần tóm tắt cuối mỗi notebook và kết luận ở notebook 04. Nếu thay đổi dữ liệu hoặc tham số, phải chạy lại rồi sửa các nhận xét cho khớp.
2. **Hằng số được khai báo lặp lại ở đầu từng notebook.** `REF_DATE` có ở cả 4 notebook, còn `ACTIVE_START` và `PRED_END` chỉ có ở notebook 01. Muốn đổi mốc thời gian phải đổi ở tất cả các nơi.
3. **Chạy lại từ notebook bị thay đổi trở đi:**

| Nếu thay đổi | Chạy lại |
|---|---|
| Dữ liệu gốc, cách làm sạch, mốc thời gian, định nghĩa churn | 01 -> 02 -> 03 -> 04 |
| Cách tính RFM, số cụm, quy tắc đặt tên cụm | 02 -> 03 -> 04 |
| Đặc trưng hoặc mô hình churn | 03 -> 04 |
| Tham số luật kết hợp, cách chọn khách nguy cơ cao, cách gợi ý | 04 |

4. **Không dùng dữ liệu từ 10/09/2011 trở đi** để tính đặc trưng hoặc làm sạch dữ liệu giai đoạn quan sát.

---

## 9. Vấn đề đã biết

- **SHAP và XGBoost:** `shap` 0.49.1 không đọc được mô hình của XGBoost 3.2 (báo lỗi `ValueError: could not convert string to float`). Vì vậy notebook 03 tính giá trị SHAP bằng hàm có sẵn trong XGBoost (`pred_contribs=True`, cũng là thuật toán TreeSHAP) và chỉ dùng `shap` để vẽ. Nếu mô hình tốt nhất không phải XGBoost, code tự chuyển sang `shap.TreeExplainer`.
- **numpy 2:** chưa chạy thử. Những chỗ in nhãn cụm và kích thước cụm đã được ép về `int` hoặc list Python thuần để output không hiện `np.int64(...)`.
- **Thư viện thừa:** `seaborn` và `imbalanced-learn` có trong `requirements.txt` nhưng chưa được dùng.
- **Quản lý phiên bản:** thư mục dự án chưa phải git repository.

---

## 10. Việc còn lại: báo cáo và slide

Báo cáo cuối cùng dự kiến gồm 8 chương. Nội dung và hình có thể lấy từ các notebook như sau:

| Chương | Lấy từ |
|---|---|
| 1. Giới thiệu và mục tiêu | Mục 1 của tài liệu này, phần mở đầu của từng notebook |
| 2. Cơ sở lý thuyết: RFM, K-Means, các thuật toán phân lớp, FP-Growth | Cần viết thêm. Các notebook có giải thích ngắn về Silhouette, Davies-Bouldin, support, confidence, lift và lý do chọn FP-Growth |
| 3. Dữ liệu và tiền xử lý | Notebook 01, hình `01_*` |
| 4. Phân cụm khách hàng | Notebook 02, hình `02_*` |
| 5. Dự đoán khách hàng rời bỏ | Notebook 03, hình `03_*` |
| 6. Luật kết hợp và gợi ý sản phẩm | Notebook 04 (mục 2 đến 5), hình `04_*` |
| 7. Đề xuất kinh doanh | Notebook 04, mục 6 |
| 8. Kết luận và hướng phát triển | Notebook 04, mục 8 |

**Hạn chế chính** (đã ghi ở notebook 04, mục 8):

- Mô hình chỉ có dữ liệu giao dịch, nên ROC-AUC dừng ở mức khá (0,73 đến 0,77).
- Chỉ có một mốc chia thời gian, và giai đoạn dự đoán rơi vào mùa cao điểm.
- Nhiều luật kết hợp chỉ là các biến thể màu của cùng một sản phẩm.
- Gợi ý mới được đánh giá trên hành vi mua tự nhiên, chưa có thử nghiệm thực tế.

**Hướng phát triển:**

- Thêm đặc trưng: tốc độ mua, xu hướng chi tiêu gần đây, khoảng cách giữa hai lần mua, tỷ lệ hủy đơn.
- Kiểm định mô hình theo nhiều mốc thời gian (*backtesting*).
- Chọn ngưỡng phân lớp theo chi phí và lợi ích kinh doanh.
- Gom sản phẩm theo danh mục để tìm luật tổng quát hơn.
- Thử nghiệm A/B cho chiến dịch giữ chân khách.
