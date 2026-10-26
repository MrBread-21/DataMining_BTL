# Bàn giao dự án: Phân khúc khách hàng, dự đoán rời bỏ và gợi ý sản phẩm

Đồ án môn **Khai phá dữ liệu**. Tài liệu này giúp người nhận bàn giao nắm được dự án làm gì, mỗi bước được làm như thế nào và chạy lại code ra sao.

- **Trạng thái:** cả 4 notebook đã hoàn chỉnh và chạy được từ đầu đến cuối. Output, 21 biểu đồ và 6 file dữ liệu trung gian đã có sẵn.
- **Phạm vi kiến thức:** dự án **chỉ dùng kiến thức trong 6 chương bài giảng** (C1, C2, C3, C4, C5, C7). Bảng đối chiếu từng bước với chương bài giảng ở [mục 11](#11-đối-chiếu-kiến-thức-với-bài-giảng).
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
| Phân cụm | R, F, M rời rạc hóa thành điểm 1 đến 5; K-Means, k = 4, ra 4 cụm: **VIP, Khách mới, Có nguy cơ, Ngủ đông**. VIP chiếm 30% số khách nhưng tạo ra 77% doanh thu. Tỷ lệ churn: VIP 21%, Có nguy cơ 53%, Khách mới 54%, Ngủ đông 74% |
| Dự đoán churn | Random Forest được chọn theo cross-validation: accuracy 0,703 trên CV, 0,678 trên tập test (mốc đoán lớp đông hơn: 0,508). Monetary, Recency, Frequency và nhãn cụm là các yếu tố quan trọng nhất; khoảng 4 tháng (128 ngày) không mua là ngưỡng nên cảnh báo |
| Luật kết hợp và gợi ý | 1.489 luật theo cụm (FP-Growth và Apriori cho cùng tập phổ biến) và 1.387 luật mua kèm toàn cửa hàng, phần lớn có dạng "mua đủ bộ, đủ màu". Có gợi ý cho 1.076 khách nguy cơ cao, mỗi gợi ý kèm lý do bằng lời; 82% khách có gợi ý dựa trên sản phẩm đã mua. Độ chính xác của gợi ý: luật của cụm 6,2%, hàng bán chạy 5,2%, luật toàn cửa hàng 4,0% |

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
|   `-- processed/               # dữ liệu trung gian do các notebook sinh ra
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

Cần **Python 3.10 trở lên**. Dự án đã chạy với Python 3.10 trên Windows 11.

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
| scikit-learn | 1.7.2 | K-Means, Decision Tree, Random Forest, AdaBoost, Gaussian Naive Bayes, cross-validation |
| scipy | 1.15.3 | Phân cụm phân cấp và dendrogram (notebook 02), kiểm định Chi-bình phương (notebook 03), ma trận thưa để đếm cặp sản phẩm (notebook 04) |
| mlxtend | 0.23.4 | FP-Growth, Apriori, sinh luật. Bắt buộc từ 0.23.4, vì `association_rules()` cần tham số `num_itemsets` |
| matplotlib | 3.10.8 | |
| nbconvert | 7.17 | Cài kèm gói `jupyter` |

`openpyxl` chỉ cần khi đọc file Excel gốc.

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
| 03 Dự đoán churn | khoảng 1 phút |
| 04 Luật kết hợp | khoảng 1 phút |
| **Tổng** | **khoảng 3 đến 4 phút** |

Khi chạy bằng dòng lệnh trên Windows, terminal có thể in `RuntimeWarning: Proactor event loop does not implement add_reader...` và, ở notebook 03, `KeyError: '...joblib_memmapping_folder...'`. Các cảnh báo này đến từ thư viện giao tiếp của Jupyter và từ bước dọn thư mục tạm của joblib khi chạy song song, không ảnh hưởng kết quả và không xuất hiện trong notebook.

### Kết quả có lặp lại được không?

Có. Mọi chỗ có yếu tố ngẫu nhiên đều cố định `RANDOM_STATE = 42`. Ở notebook 04, các luật và sản phẩm bằng điểm nhau được xếp tiếp theo tiêu chí phụ (confidence -> support -> tên sản phẩm). Đã kiểm tra: chạy lại cả 4 notebook hai lần cho cùng số liệu và cùng các file CSV. Ngoại lệ duy nhất là cột **thời gian chạy** trong bảng so sánh Apriori và FP-Growth (notebook 04, mục 2.1), vốn phụ thuộc vào máy và thay đổi nhẹ giữa các lần chạy.

---

## 6. Cách làm

### 6.1. Định nghĩa rời bỏ (churn) và các mốc thời gian

Đây là phần cốt lõi của dự án, tương ứng với bước **chọn lựa dữ liệu** của quy trình KDD (C1) và bước **xác định thuộc tính lớp** của bài toán phân lớp (C4). Thay đổi bất kỳ mốc nào ở đây đều làm thay đổi toàn bộ kết quả phía sau.

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

**Các bước làm sạch** (notebook in số dòng còn lại sau mỗi bước), mỗi bước gắn với một nhóm vấn đề chất lượng dữ liệu trong C2:

| Bước | Nội dung | Nhóm vấn đề (C2) | Số dòng còn lại |
|---|---|---|---|
| 1 | Bỏ dòng trùng lặp (34.335 dòng, trong đó 22.844 dòng do hai sheet chồng lấn nhau từ 01/12 đến 09/12/2010) | Tích hợp, dư thừa | 1.033.036 |
| 2 | Bỏ dòng thiếu `Customer ID`, đổi tên cột thành `CustomerID`, ép về kiểu số nguyên | Dữ liệu thiếu (bỏ qua bản ghi) | 797.885 |
| 3a | Bỏ đơn mua gốc đã bị hủy (xem giải thích bên dưới) | Không nhất quán | 791.852 |
| 3b | Bỏ hóa đơn hủy (`C`) và hóa đơn điều chỉnh (`A`) | Không nhất quán | 773.462 |
| 4 | Bỏ dòng có `Quantity <= 0` hoặc `Price <= 0` | Nhiễu | 773.393 |
| 5 | Chỉ giữ mã sản phẩm thật: `StockCode` bắt đầu bằng 5 chữ số (regex `^\d{5}`), loại phí ship, giảm giá, điều chỉnh thủ công... | Nhiễu | 770.688 |
| 6 | Chuẩn hóa tên sản phẩm: bỏ khoảng trắng thừa, mỗi `StockCode` dùng tên xuất hiện nhiều nhất | Không nhất quán | 770.688 |
| 7 | Tạo cột `TotalPrice = Quantity * Price` | Xây dựng thuộc tính | 770.688 |

**Bước 3a** không có trong cách làm sạch thông thường nhưng cần thiết:

- **Vấn đề:** nếu chỉ bỏ dòng hủy thì đơn mua gốc vẫn còn, làm doanh thu bị thổi phồng. Ví dụ: một đơn 80.995 sản phẩm (£168K) bị hủy ngay trong ngày.
- **Cách làm:** với mỗi dòng hủy, tìm dòng mua gần nhất trước đó của cùng khách, cùng sản phẩm, cùng số lượng (dùng `pd.merge_asof`), rồi loại dòng mua đó.
- **Điều kiện:** chỉ ghép khi đơn hủy và đơn gốc nằm **cùng một giai đoạn** so với `REF_DATE`. Nếu không, ta đã dùng thông tin tương lai để sửa dữ liệu quá khứ. 331 dòng hủy ở giai đoạn dự đoán có đơn gốc ở giai đoạn quan sát nên không được ghép.
- **Kết quả:** loại 6.033 dòng mua, khoảng £666K doanh thu ảo.

**Bước 6** giúp luật kết hợp ở notebook 04 không bị tách đôi chỉ vì cùng một sản phẩm có nhiều cách ghi tên. Sau bước này còn 4.611 mã sản phẩm với 4.574 tên.

**EDA:** doanh thu theo tháng (tính mùa vụ rõ, đỉnh vào tháng 11), top 10 sản phẩm, top 10 quốc gia (Anh chiếm 83,5% doanh thu), histogram số hóa đơn mỗi khách, phân phối giá trị đơn hàng.

**Thống kê mô tả giá trị đơn hàng (C2):** trung bình £453, trung vị £301, mode £15, midrange £21.814, Q1 £158, Q3 £471, IQR £313, độ lệch chuẩn £914. 2.830 hóa đơn (7,8%) vượt Q3 + 1,5×IQR (£939), chiếm 38,1% doanh thu. Các hóa đơn này là đơn bán sỉ thật (trung bình 41 loại sản phẩm, 1.411 đơn vị hàng) nên được **giữ lại**; trung vị được dùng để mô tả hóa đơn điển hình vì không bị ngoại lai ảnh hưởng.

### 6.3. Notebook 02: Phân cụm RFM

1. **Tính RFM** cho mỗi khách từ giao dịch giai đoạn quan sát (gộp nhóm theo khách, C2):
   - `Recency`: số ngày từ lần mua cuối đến 10/09/2011.
   - `Frequency`: số hóa đơn khác nhau.
   - `Monetary`: tổng `TotalPrice`.
2. **Thống kê mô tả (C2):** F và M lệch phải mạnh; theo quy tắc 1,5×IQR có 399 khách ngoại lai về Frequency và 435 khách ngoại lai về Monetary (khoảng 10%), nắm 43% tổng số hóa đơn và 61% tổng chi tiêu. K-Means trên giá trị gốc sẽ bị vài khách cực lớn chi phối tâm cụm.
3. **Rời rạc hóa (C2, đóng thùng equal-frequency):** mỗi chỉ số được chia thành 5 thùng có số khách bằng nhau, gán điểm 1 đến 5: `pd.qcut(s.rank(method="first"), 5, labels=False) + 1`. `rank(method="first")` phá hòa vì nhiều khách có cùng Frequency. Recency đảo chiều (`R_score = 6 - q`), nên với cả ba điểm, cao là tốt. Ba điểm cùng thang 1 đến 5 nên không cần chuẩn hóa thêm.
4. **Chọn k (C7):** thử k = 2..10 trên ma trận `[R_score, F_score, M_score]`, đo Inertia (hàm mục tiêu P(W, Q) của K-Means), **độ chặt** (khoảng cách Euclid trung bình tới tâm cụm) và **độ tách** (khoảng cách nhỏ nhất giữa hai tâm cụm). Chọn **k = 4**: inertia giảm 2.507 khi từ 3 lên 4 nhưng chỉ giảm 894 khi từ 4 lên 5 (điểm gãy Elbow rõ nhất); độ tách 2,469 cao nhất trong các k ≥ 3; độ chặt 1,157.
5. **So sánh thuật toán (C7):** phân cụm phân cấp với 4 kiểu liên kết, cắt thành 4 cụm:

| Phương án | Kích thước các cụm | Độ chặt | Độ tách |
|---|---|---|---|
| **K-Means** | 1.284, 1.121, 947, 947 | **1,157** | **2,469** |
| Liên kết single | chỉ 1 cụm (4.299) | 2,307 | không xác định |
| Liên kết complete | 1.472, 1.412, 764, 651 | 1,272 | 2,284 |
| Liên kết average | 2.838, 1.458, 2, 1 | 1,621 | 3,412 |
| Liên kết centroid | 2.963, 1.334, 1, 1 | 1,643 | 3,394 |

   Ma trận điểm chỉ có 110 điểm khác nhau xếp thành lưới bước 1, nên liên kết single gộp mọi thứ ở cùng khoảng cách 1 (hiện tượng "chuỗi dài" cực đoan). Average và centroid tạo ra cụm 1 đến 2 khách. Complete tạo cụm gọn nhưng kém K-Means ở cả độ chặt và độ tách. **K-Means được chọn.** Dendrogram của liên kết complete được vẽ ở `02_dendrogram.png`.
6. **Đặt tên cụm** bằng quy tắc viết thành code, tính trên giá trị gốc trung bình của R, F, M, để kết quả không phụ thuộc vào thứ tự nhãn ngẫu nhiên của K-Means:
   - Cụm có Frequency trung bình cao nhất là **VIP**.
   - Trong các cụm còn lại, cụm có Recency thấp nhất là **Khách mới**.
   - Trong hai cụm còn lại, cụm có Frequency cao hơn là **Có nguy cơ**, cụm kia là **Ngủ đông**.
   - Cột `Cluster` được đánh số lại: 0 = VIP, 1 = Khách mới, 2 = Có nguy cơ, 3 = Ngủ đông (từ nhóm giá trị cao nhất đến thấp nhất).

| Cụm | Số khách | % doanh thu | Tỷ lệ churn | Recency TB (ngày) | Frequency TB | Monetary TB | Điểm R, F, M TB |
|---|---|---|---|---|---|---|---|
| VIP | 1.284 | 76,7% | 21,3% | 38 | 14,9 | £7.605 | 4,4 / 4,6 / 4,5 |
| Khách mới | 947 | 4,0% | 53,9% | 67 | 2,1 | £535 | 3,9 / 2,0 / 1,9 |
| Có nguy cơ | 947 | 15,7% | 52,8% | 206 | 5,5 | £2.108 | 2,1 / 3,6 / 3,6 |
| Ngủ đông | 1.121 | 3,6% | 74,2% | 278 | 1,6 | £413 | 1,5 / 1,6 / 1,7 |

7. **Diễn giải cụm (C7):** boxplot R, F, M gốc theo cụm; điểm R, F, M trung bình theo cụm (`02_snake_plot.png`); scatter plot Recency × Frequency (`02_scatter_rf.png`); và **cây quyết định mô tả cụm** (độ sâu 3, đầu vào R, F, M gốc) xếp đúng cụm cho 91,3% khách. Luật tiêu biểu: Recency ≤ 78 ngày và Frequency > 4 -> VIP; Recency > 158 ngày, Monetary ≤ £717 và Frequency ≤ 4 -> Ngủ đông.

### 6.4. Notebook 03: Dự đoán churn

**Đặc trưng** (chỉ từ giai đoạn quan sát, giá trị gốc, không biến đổi):

- `Recency`, `Frequency`, `Monetary`.
- `Tenure`: số ngày từ lần mua đầu tiên đến 10/09/2011.
- `AvgOrderValue`: giá trị trung bình mỗi hóa đơn.
- `NumProducts`: số sản phẩm khác nhau đã mua.
- `IsUK`: 1 nếu khách ở Anh, 0 nếu ở nước khác.
- `Cluster`: nhãn cụm từ notebook 02, dùng trực tiếp số nguyên 0 đến 3.

**Phân tích thuộc tính (C2):**

- Tương quan Pearson cao nhất: Frequency và NumProducts (0,69), Frequency và Monetary (0,63). Không cặp nào vượt 0,8 nên không bỏ thuộc tính nào (`03_tuong_quan.png`).
- Kiểm định Chi-bình phương (α = 0,05): `IsUK` × `churn` có χ² = 0,93 < 3,841, **độc lập**; `ClusterName` × `churn` có χ² = 692,6 > 7,815, **phụ thuộc**.

**Huấn luyện:**

- Chia train/test theo tỷ lệ 80/20 bằng lấy mẫu phân tầng (`stratify=churn`, 3.439 / 860 khách).
- Bốn mô hình theo bài giảng: Decision Tree, Random Forest, AdaBoost (C4) và Gaussian Naive Bayes (C5).
- Chọn tham số bằng vòng lặp tường minh với `cross_val_score(..., scoring="accuracy")` trên 5-fold `StratifiedKFold` của tập train. Decision Tree thử `criterion` ∈ {gini, entropy}, `max_depth` ∈ {3, 4, 5, 6, 8}, `min_samples_leaf` ∈ {10, 30, 50} (điều kiện dừng của cây); Random Forest 300 cây, `max_features="sqrt"`, `min_samples_leaf` ∈ {5, 20}; AdaBoost 200 cây một tầng.
- Tiêu chí: **accuracy**, phù hợp vì hai lớp gần cân bằng (49,2% churn).
- **Mô hình được chọn theo accuracy CV, không theo điểm test.**

| Mô hình | Tham số tốt nhất | Accuracy CV | Accuracy test |
|---|---|---|---|
| Decision Tree | gini, `max_depth=6`, `min_samples_leaf=50` | 0,690 ± 0,003 | 0,658 |
| **Random Forest** (được chọn) | `min_samples_leaf=20` | **0,703 ± 0,007** | **0,678** |
| AdaBoost | 200 cây một tầng | 0,696 ± 0,008 | 0,664 |
| Gaussian NB | | 0,630 ± 0,012 | 0,608 |

Mốc so sánh (đoán mọi khách là ở lại): 0,508. Trong từng cụm, mô hình chỉ hơn cách đoán theo lớp đông hơn một chút, trừ Khách mới (0,636 so với 0,518).

**Giải thích mô hình:**

- **Cây quyết định -> tập luật (C4):** tầng đầu tách theo nhãn cụm (VIP với phần còn lại, rồi Ngủ đông với Khách mới và Có nguy cơ). Ngưỡng Recency khoảng 4 tháng vẫn xuất hiện: trong nhóm Khách mới và Có nguy cơ, khách mua 13 đến 108 loại sản phẩm, chi không quá £1.628 và đã hơn 128 ngày không mua có tỷ lệ rời bỏ 69,7%, so với 41,8% nếu còn mua trong 128 ngày. Notebook in toàn bộ tập luật (`export_text`) và bảng luật của 35 lá.
- **Độ quan trọng theo mức giảm Gini (Random Forest):** Monetary 0,216, Recency 0,184, Frequency 0,163, Cluster 0,152, NumProducts 0,123, Tenure 0,093, AvgOrderValue 0,069, IsUK 0,001.

**Dự đoán cho toàn bộ khách:**

- Khách thuộc tập train nhận xác suất *out-of-fold* (`cross_val_predict` với cùng 5 fold): mỗi khách được dự đoán bởi mô hình huấn luyện trên D \ Dᵢ, không chứa fold của khách đó.
- Khách thuộc tập test nhận dự đoán thông thường. Với Random Forest, xác suất churn là tỷ lệ phiếu của 300 cây.
- Như vậy không xác suất nào được tính trên chính dữ liệu mà mô hình đã học.

**Lưu ý quan trọng:** ở ngưỡng 0,5, mô hình chỉ cảnh báo được 17 khách VIP, dù thực tế có 274 khách VIP rời bỏ (xác suất của khách VIP đều thấp). Tuy vậy, bên trong cụm VIP mô hình vẫn **xếp hạng** khách khá tốt: trên tập test, 25% khách VIP có xác suất cao nhất rời bỏ 47,7%, gần gấp đôi mức chung 24,1%. Đây là lý do notebook 04 chọn khách theo thứ hạng trong từng cụm.

### 6.5. Notebook 04: Luật kết hợp và gợi ý sản phẩm

**Khai phá luật theo từng cụm (C3):**

1. Lấy giao dịch giai đoạn quan sát và gắn nhãn cụm của khách.
2. Tạo ma trận giỏ hàng: mỗi dòng là một hóa đơn, mỗi cột là một sản phẩm (`Description`).
3. Chọn `min_support = max(1%, 20 / số hóa đơn của cụm)`, để cụm nhỏ vẫn cần ít nhất 20 hóa đơn cho một luật.
4. Bỏ trước các sản phẩm có support thấp hơn `min_support` (theo tính chất Apriori), nhưng vẫn giữ đủ mọi hóa đơn để support được tính đúng mẫu số.
5. Chạy **FP-Growth** (`mlxtend`), sinh luật bằng `association_rules(..., metric="confidence", min_threshold=0.5)`. Mọi sản phẩm đều có support nhỏ hơn 0,5 (cao nhất 18,2%), nên mọi luật có confidence ≥ 0,5 đều thỏa P(B | A) > P(B).
6. Sắp xếp luật theo confidence, rồi support, rồi tên sản phẩm.

Kết quả: VIP 520 luật, Khách mới 117, Có nguy cơ 729, Ngủ đông 123 (tổng 1.489). 69,0% luật của VIP và 76,1% luật của Có nguy cơ chỉ có ở cụm đó; VIP và Có nguy cơ có 147 luật chung (bảng đếm luật chung 4×4 trong notebook).

**So sánh Apriori và FP-Growth (C3):** trên cùng ma trận giỏ hàng, hai thuật toán cho **cùng tập phổ biến và cùng support** ở cả 4 cụm. Trong `mlxtend`, Apriori (`low_memory=True`) lại nhanh hơn FP-Growth khoảng 13 đến 34 lần, vì sau khi lọc sản phẩm hiếm số ứng viên nhỏ và Apriori được cài đặt bằng phép toán vector, còn FP-Growth bằng Python thuần.

**Cảnh báo:** các luật confidence 100% của cụm Có nguy cơ (bộ đồ gỗ trắng WOODEN FRAME, WOOD CABINET, EDWARDIAN PARASOL) gần như chỉ đến từ **một khách** (17850, chiếm 98,5% số hóa đơn của luật mạnh nhất). Support tính theo hóa đơn nên một khách đặt lặp lại có thể tự tạo ra cả một nhóm luật.

**Chọn khách nguy cơ cao:** 25% khách có xác suất churn cao nhất **trong từng cụm** (từ phân vị 75%, C2), tổng cộng 1.076 khách. Nhóm được chọn có tỷ lệ churn thực tế cao hơn rõ rệt so với mức chung của cụm (ví dụ VIP: 43,0% so với 21,3%). 321 khách VIP nguy cơ cao đã chi £680K, nhiều hơn 3 cụm còn lại cộng lại.

**Luật mua kèm toàn cửa hàng** (dùng khi luật của cụm không đủ):

- **Vấn đề:** luật của cụm cần ít nhất 1% số hóa đơn của cụm, nên nhiều sản phẩm ít phổ biến không có luật nào. Nếu chỉ dùng luật của cụm, 473 trên 1.076 khách nguy cơ cao (44%) không nhận được gợi ý nào dựa trên sản phẩm đã mua.
- **Cách làm:** tìm thêm luật một sản phẩm -> một sản phẩm trên toàn bộ hóa đơn giai đoạn quan sát (gộp mọi cụm), với cùng tiêu chuẩn: ít nhất 20 hóa đơn chứa cả hai sản phẩm và confidence ≥ 0,5.
- **Cách đếm:** đây là bước đếm support của các 2-itemset ứng viên trong Apriori (C₂ -> L₂, C3). Notebook đếm số hóa đơn chứa từng cặp sản phẩm bằng phép nhân ma trận thưa (`basket.T @ basket`, dùng `scipy.sparse`), chỉ mất vài giây. Với ngưỡng thấp như vậy, FP-Growth chạy hơn 10 phút chưa xong.
- **Kết quả:** 1.387 luật, vế trái gồm 839 sản phẩm (luật theo cụm chỉ có 192 sản phẩm ở vế trái).

**Thuật toán gợi ý** (5 sản phẩm mỗi khách, bán chéo), điền lần lượt từ 3 nguồn:

1. **Luật của cụm** mà khách thuộc về, theo thứ tự confidence giảm dần. Nếu mọi sản phẩm ở vế trái của luật khách đều đã mua, gợi ý các sản phẩm ở vế phải mà khách chưa mua.
2. Nếu chưa đủ 5 sản phẩm, dùng tiếp **luật toàn cửa hàng** theo cùng cách.
3. Nếu vẫn chưa đủ, bổ sung bằng **hàng bán chạy nhất trong cụm** mà khách chưa mua. Đây là gợi ý theo độ phổ biến, không dựa trên sản phẩm khách đã mua.

Các cột của bảng gợi ý:

- Cột `Source` ghi nguồn: `Luật của cụm`, `Luật toàn cửa hàng` hoặc `Bán chạy trong cụm`.
- Cột `BasedOn` ghi sản phẩm khách đã mua làm căn cứ.
- Cột `Reason` giải thích lý do của từng gợi ý bằng lời, do các hàm `explain_rule()` và `fallback_reason()` tạo ra:
  - **Gợi ý từ luật:** nêu sản phẩm khách đã mua, luật đó là của cụm hay của toàn cửa hàng, tỷ lệ hóa đơn có các sản phẩm đó cũng có sản phẩm được gợi ý (confidence), số hóa đơn chứa đồng thời các sản phẩm và tỷ lệ trên tổng số hóa đơn (support).
  - **Gợi ý từ hàng bán chạy:** nêu cụ thể khách đã mua gì. Code xác định đúng một trong ba trường hợp: không luật nào khớp với sản phẩm đã mua; có luật khớp nhưng luật chỉ gợi ý sản phẩm khách đã có; hoặc luật chỉ cho được vài gợi ý. Sau đó nêu thứ hạng và độ phổ biến của sản phẩm trong cụm.

**Kết quả gợi ý:** 82% khách nguy cơ cao có ít nhất 1 gợi ý dựa trên sản phẩm đã mua. 195 khách (18%) chỉ nhận hàng bán chạy; phần lớn họ thuộc cụm Ngủ đông (85) và Khách mới (72).

**Độ chính xác của gợi ý:**

- So sánh gợi ý với những gì khách thực sự mua trong giai đoạn dự đoán (chỉ để đánh giá). 366 trên 1.076 khách nguy cơ cao đã quay lại mua.
- Độ chính xác: luật của cụm 6,2%, hàng bán chạy 5,2%, luật toàn cửa hàng 4,0%.
- Luật của cụm có độ chính xác cao nhất ở VIP (6,8% so với 1,4% của hàng bán chạy), Khách mới và Có nguy cơ. Ngoại lệ là cụm Ngủ đông: ở đây hàng bán chạy cao nhất (10,2%).

Cuối notebook 04 có phần **đề xuất chiến lược cho từng cụm** và **kết luận chung** của dự án (kết quả, hạn chế, hướng phát triển). Mỗi cụm được trình bày theo 4 phần: tình hình, hành động đề xuất, căn cứ (số liệu từ các notebook giải thích vì sao đề xuất như vậy) và sản phẩm gợi ý.

Mọi bảng và biểu đồ trong notebook 04 đều có phần **Cách đọc** giải thích ý nghĩa các cột, trước phần **Nhận xét**.

---

## 7. Mô tả các file đầu ra

Các file có tiếng Việt được ghi với encoding `utf-8-sig`, nên mở bằng Excel không bị lỗi font.

| File | Kích thước | Các cột |
|---|---|---|
| `transactions_clean.csv` | 770.688 × 9 | Giao dịch đã làm sạch: `Invoice`, `StockCode`, `Description`, `Quantity`, `InvoiceDate`, `Price`, `CustomerID`, `Country`, `TotalPrice` (toàn bộ thời gian, gồm cả giai đoạn dự đoán) |
| `customers.csv` | 4.299 × 2 | `CustomerID`, `churn` |
| `customers_clustered.csv` | 4.299 × 10 | `CustomerID`, `Recency`, `Frequency`, `Monetary`, `R_score`, `F_score`, `M_score` (điểm 1 đến 5), `Cluster` (0 đến 3), `ClusterName`, `churn` |
| `churn_predictions.csv` | 4.299 × 6 | `CustomerID`, `ClusterName`, `churn_proba` (xác suất rời bỏ), `churn_pred` (nhãn dự đoán ở ngưỡng 0,5), `churn` (nhãn thật), `Set` (`train` hoặc `test`) |
| `rules_by_cluster.csv` | 1.489 × 7 | `ClusterName`, `Antecedents`, `Consequents` (nhiều sản phẩm nối bằng ` + `), `AntecedentSupport`, `ConsequentSupport`, `Support`, `Confidence`; sắp theo cụm rồi theo confidence giảm dần |
| `recommendations.csv` | 5.380 × 10 | `CustomerID`, `ClusterName`, `churn_proba`, `Rank` (1 đến 5), `RecommendedProduct`, `Source` (`Luật của cụm`, `Luật toàn cửa hàng` hoặc `Bán chạy trong cụm`), `BasedOn`, `Confidence`, `Support` (ba cột này để trống nếu là hàng bán chạy), `Reason` (lý do gợi ý viết bằng lời) |

Mỗi biểu đồ trong `figures/` có tiền tố là số của notebook sinh ra nó (ví dụ `01_doanh_thu_theo_thang.png`, `03_so_sanh_mo_hinh.png`). Cả 21 hình đều do notebook sinh ra:

| Notebook | Hình |
|---|---|
| 01 | `01_doanh_thu_theo_thang`, `01_top10_san_pham`, `01_top10_quoc_gia`, `01_phan_phoi_so_hoa_don`, `01_phan_phoi_gia_tri_don_hang`, `01_ty_le_churn` |
| 02 | `02_phan_phoi_rfm`, `02_chon_so_cum_k`, `02_dendrogram`, `02_snake_plot`, `02_boxplot_rfm_theo_cum`, `02_scatter_rf`, `02_ty_trong_va_churn_theo_cum`, `02_cay_mo_ta_cum` |
| 03 | `03_dac_trung_theo_churn`, `03_tuong_quan`, `03_so_sanh_mo_hinh`, `03_cay_quyet_dinh`, `03_feature_importance`, `03_xac_suat_churn_theo_cum` |
| 04 | `04_luat_theo_cum` |

---

## 8. Quy ước và lưu ý khi chỉnh sửa

### Quy ước

- Đường dẫn trong code là đường dẫn tương đối, tính từ thư mục `notebooks/`.
- `RANDOM_STATE = 42` ở mọi chỗ có yếu tố ngẫu nhiên.
- Ô markdown viết bằng tiếng Việt, gồm mục đích và nhận xét kết quả của từng bước, có thể chép thẳng vào báo cáo. Khi một bước dùng kiến thức của bài giảng, markdown ghi ngắn gọn công thức hoặc định nghĩa và nêu chương (ví dụ "(C7, đánh giá cụm)"). Tên biến và tên hàm viết bằng tiếng Anh. Code không có chú thích (comment); mọi giải thích nằm trong các ô markdown.
- Chỉ dùng kiến thức trong 6 chương bài giảng. Kỹ thuật ngoài bài giảng chỉ được nhắc ở mục "Hướng phát triển" cuối notebook 04.
- Biểu đồ được lưu bằng hàm `save_fig()` khai báo ở đầu mỗi notebook. Trục log chỉ dùng để hiển thị và được ghi rõ trên tiêu đề hoặc nhãn trục; dữ liệu không bị biến đổi.
- Màu của mỗi cụm cố định trong mọi biểu đồ: VIP xanh dương, Khách mới cam, Có nguy cơ xanh lá, Ngủ đông vàng. Nếu thêm biểu đồ mới, nên giữ đúng màu này.

### Lưu ý khi chỉnh sửa

1. **Các con số trong ô nhận xét được viết tay** theo kết quả hiện tại, kể cả phần tóm tắt cuối mỗi notebook và kết luận ở notebook 04. Nếu thay đổi dữ liệu hoặc tham số, phải chạy lại rồi sửa các nhận xét cho khớp.
2. **Hằng số được khai báo lặp lại ở đầu từng notebook.** `REF_DATE` có ở cả 4 notebook, còn `ACTIVE_START` và `PRED_END` chỉ có ở notebook 01. Muốn đổi mốc thời gian phải đổi ở tất cả các nơi.
3. **Chạy lại từ notebook bị thay đổi trở đi:**

| Nếu thay đổi | Chạy lại |
|---|---|
| Dữ liệu gốc, cách làm sạch, mốc thời gian, định nghĩa churn | 01 -> 02 -> 03 -> 04 |
| Cách tính RFM, cách rời rạc hóa, số cụm, quy tắc đặt tên cụm | 02 -> 03 -> 04 |
| Đặc trưng hoặc mô hình churn | 03 -> 04 |
| Tham số luật kết hợp, cách chọn khách nguy cơ cao, cách gợi ý | 04 |

4. **Không dùng dữ liệu từ 10/09/2011 trở đi** để tính đặc trưng hoặc làm sạch dữ liệu giai đoạn quan sát.

---

## 9. Vấn đề đã biết

- **Liên kết single không cắt được thành 4 cụm:** ma trận điểm R, F, M chỉ có 110 điểm khác nhau xếp thành lưới bước 1, nên mọi lần gộp đều ở khoảng cách 1 và `fcluster` trả về 1 cụm. Hàm `compactness_separation()` trả về độ tách rỗng (NaN) trong trường hợp này; bảng hiển thị "không xác định".
- **Thời gian Apriori và FP-Growth:** thay đổi nhẹ giữa các lần chạy, và trong `mlxtend` FP-Growth chậm hơn Apriori. Nhận xét trong notebook chỉ nêu tương quan, không dựa vào con số chính xác.
- **Luật do một khách tạo ra:** khách 17850 (cụm Có nguy cơ) đặt lặp lại cùng một bộ đồ gỗ trắng, tạo ra phần lớn các luật confidence 100% của cụm. Nên loại các luật này trước khi dùng cho chiến dịch thật.
- **numpy 2:** chưa chạy thử. Những chỗ in nhãn cụm và kích thước cụm đã được ép về `int` hoặc list Python thuần để output không hiện `np.int64(...)`.

---

## 10. Việc còn lại: báo cáo và slide

Báo cáo cuối cùng dự kiến gồm 8 chương. Nội dung và hình có thể lấy từ các notebook như sau:

| Chương | Lấy từ |
|---|---|
| 1. Giới thiệu và mục tiêu | Mục 1 của tài liệu này, phần mở đầu của từng notebook |
| 2. Cơ sở lý thuyết | Dùng bảng ở [mục 11](#11-đối-chiếu-kiến-thức-với-bài-giảng) làm khung. Các ô markdown của notebook đã ghi sẵn công thức và định nghĩa kèm chương: thống kê mô tả, 1,5×IQR, rời rạc hóa, Pearson, Chi-bình phương (C2); hàm mục tiêu K-Means, độ chặt, độ tách, các kiểu liên kết (C7); Gini, entropy, điều kiện dừng, bagging, Random Forest, AdaBoost, K-fold (C4); Naive Bayes (C5); support, confidence, tính chất Apriori, FP-Growth (C3) |
| 3. Dữ liệu và tiền xử lý | Notebook 01, hình `01_*` |
| 4. Phân cụm khách hàng | Notebook 02, hình `02_*` |
| 5. Dự đoán khách hàng rời bỏ | Notebook 03, hình `03_*` |
| 6. Luật kết hợp và gợi ý sản phẩm | Notebook 04 (mục 2 đến 5), hình `04_*` |
| 7. Đề xuất kinh doanh | Notebook 04, mục 6 |
| 8. Kết luận và hướng phát triển | Notebook 04, mục 8 |

**Hạn chế chính** (đã ghi ở notebook 04, mục 8):

- Mô hình chỉ có dữ liệu giao dịch, nên độ chính xác dừng ở mức khá (khoảng 0,68 đến 0,70).
- Chỉ có một mốc chia thời gian, và giai đoạn dự đoán rơi vào mùa cao điểm.
- Support tính theo hóa đơn, nên một khách đặt lặp lại có thể tạo ra cả nhóm luật; nhiều luật chỉ là các biến thể màu của cùng một sản phẩm.
- Gợi ý mới được đánh giá trên hành vi mua tự nhiên, chưa có thử nghiệm thực tế.

**Hướng phát triển** (kỹ thuật ngoài bài giảng chỉ nêu ở đây):

- Thêm đặc trưng: tốc độ mua, xu hướng chi tiêu gần đây, khoảng cách giữa hai lần mua, tỷ lệ hủy đơn.
- Kiểm định mô hình theo nhiều mốc thời gian (*backtesting*).
- Thử thêm các độ đo đánh giá phân lớp khác ngoài accuracy và chọn ngưỡng phân lớp theo chi phí và lợi ích kinh doanh.
- Thử thêm độ đo khác cho luật kết hợp và tính support theo số khách; gom sản phẩm theo danh mục để tìm luật tổng quát hơn.
- Thử nghiệm A/B cho chiến dịch giữ chân khách.

---

## 11. Đối chiếu kiến thức với bài giảng

Bảng này dùng làm khung cho chương Cơ sở lý thuyết của báo cáo. Mỗi dòng là một bước của dự án, kỹ thuật đã dùng và chương bài giảng tương ứng.

| Bước của dự án | Kỹ thuật dùng | Chương |
|---|---|---|
| Tổng thể dự án | Quy trình KDD: làm sạch, tích hợp, chọn lựa, biến đổi, khai phá, đánh giá mẫu, biểu diễn tri thức; ứng dụng phân khúc khách hàng, giữ chân khách hàng, bán chéo | C1 |
| Định nghĩa churn, hai giai đoạn quan sát / dự đoán (NB 01) | Chọn lựa dữ liệu; xác định thuộc tính lớp | C1, C4 |
| Nối hai sheet, bỏ dòng trùng (NB 01) | Tích hợp dữ liệu, loại dư thừa | C2 |
| Bỏ dòng thiếu `Customer ID` (NB 01) | Dữ liệu thiếu: bỏ qua bản ghi | C2 |
| Ghép đơn hủy với đơn gốc bằng `pd.merge_asof`, chuẩn hóa tên sản phẩm (NB 01) | Dữ liệu không nhất quán | C2 |
| Bỏ `Quantity`/`Price` ≤ 0, mã không phải sản phẩm (NB 01) | Dữ liệu nhiễu | C2 |
| Tạo `TotalPrice`, `AvgOrderValue`, `Tenure`, `NumProducts` (NB 01, 03) | Xây dựng thuộc tính | C2 |
| Thống kê giá trị đơn hàng (NB 01) và R, F, M (NB 02) | Mean, median, mode, midrange, tứ phân vị, IQR, ngoại lai theo 1,5×IQR, độ lệch chuẩn | C2 |
| Biểu đồ EDA (NB 01, 02, 03) | Histogram, boxplot, scatter plot, biểu đồ cột | C1, C2 |
| Tính RFM theo khách (NB 02) | Gộp nhóm (aggregation) | C2 |
| Điểm R, F, M 1 đến 5 (NB 02) | Rời rạc hóa, đóng thùng equal-frequency; mục đích của chuẩn hóa | C2 |
| Chọn k (NB 02) | Hàm mục tiêu K-Means P(W, Q) (Elbow), độ chặt, độ tách | C7 |
| Phân cụm (NB 02) | K-Means; phân cụm phân cấp gộp với liên kết single, complete, average, centroid; dendrogram; độ đo Euclid | C7 |
| Diễn giải cụm (NB 02) | Trung bình cụm, boxplot, scatter plot, cây quyết định chạy trên dữ liệu đã phân cụm | C7 |
| Phân tích thuộc tính (NB 03) | Hệ số tương quan Pearson (thuộc tính dư thừa); kiểm định Chi-bình phương | C2 |
| Chia train/test (NB 03) | Lấy mẫu phân tầng (stratified) | C4 |
| Chọn tham số và mô hình (NB 03) | Stratified K-fold cross-validation; accuracy; overfitting; bias, variance | C4 |
| Mô hình phân lớp (NB 03) | Cây quyết định (CART Gini, ID3/C4.5 entropy, điều kiện dừng); Random Forest (bagging, bootstrap, √m thuộc tính); AdaBoost (boosting) | C4 |
| Mô hình Bayes (NB 03) | Định lý Bayes, Gaussian Naive Bayes, giả thiết độc lập có điều kiện | C5 |
| Giải thích mô hình (NB 03) | Cây -> tập luật; độ quan trọng theo mức giảm Gini | C4 |
| Xác suất cho toàn bộ khách (NB 03) | K-fold: mỗi khách được dự đoán bởi mô hình huấn luyện trên D \ Dᵢ; bỏ phiếu số đông của Random Forest | C4 |
| Khai phá luật theo cụm (NB 04) | Item, itemset, transaction, support, confidence, frequent itemset, tính chất Apriori, FP-Growth, Apriori | C3 |
| Luật toàn cửa hàng (NB 04) | Đếm support các 2-itemset ứng viên (C₂ -> L₂ của Apriori) | C3 |
| Chọn khách nguy cơ cao (NB 04) | Tứ phân vị (phân vị 75% trong cụm) | C2 |
| Gợi ý và đánh giá gợi ý (NB 04) | Bán chéo; đánh giá mẫu bằng độ chính xác của gợi ý | C1, C3 |
