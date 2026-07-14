# Oracle Fusion Excel Runner

MVP Streamlit app untuk mencari **reference ID** dari Oracle Fusion, membuat template Excel, generate JSON payload, validasi data, dan upload **Inventory Organizations** ke Oracle Fusion Cloud SCM via REST API.

Endpoint awal:

```http
GET  /fscmRestApi/resources/11.13.18.05/inventoryOrganizations
POST /fscmRestApi/resources/11.13.18.05/inventoryOrganizations
```

## Struktur fitur

1. **Reference Finder**
   - Input instance, username, dan password.
   - Fetch daftar ID dari LOV/reference endpoint:
     - Business Units
     - Legal Entities
     - Inventory Organizations LOV
     - Locations LOV, jika tersedia di instance
   - Download `oracle_fusion_reference_ids.xlsx`.
   - GET existing Inventory Organizations.
   - Pilih salah satu organization sebagai reference.
   - Generate template Excel, mapping JSON, sample payload JSON, dan raw GET response.

2. **Template & JSON Builder**
   - Pilih field yang mau dimasukkan ke template.
   - Download template Excel.
   - Download mapping JSON.
   - Preview sample JSON payload.

3. **Upload Runner**
   - Upload Excel/CSV.
   - Validasi kolom dan tipe data.
   - Validasi field read-only, misalnya `ManagementBusinessUnitName` tidak boleh dipakai untuk POST.
   - Preview JSON per row.
   - Test connection.
   - Test selected row.
   - Run all rows.
   - Download upload log.

## Cara menjalankan lokal

```bash
cd oracle_fusion_inventory_runner
python -m venv .venv

# Windows PowerShell
.venv\Scripts\Activate.ps1

# macOS/Linux
source .venv/bin/activate

pip install -r requirements.txt
streamlit run app.py
```

## Cara test tanpa Oracle

1. Jalankan aplikasi.
2. Buka halaman **Template & JSON Builder**.
3. Download template Excel.
4. Buka halaman **Upload Runner**.
5. Pastikan **Mock / Dry Run mode** aktif di sidebar.
6. Upload template Excel.
7. Klik **Run All Rows**.
8. Download log.

Dalam mode ini aplikasi tidak melakukan POST ke Oracle.

## Cara GET reference ID dari Oracle Fusion

Gunakan instance DEV/TEST dulu.

1. Buka halaman **Reference Finder**.
2. Isi Oracle Base URL, username, dan password.
3. Klik **Fetch Reference IDs**.
4. Download **Reference IDs Excel**.
5. Pakai ID di file itu untuk mengisi template upload.

## Cara generate template dari existing organization

1. Buka halaman **Reference Finder**.
2. Klik **GET Existing Orgs**.
3. Pilih salah satu `result_index` sebagai reference.
4. Download template Excel dan mapping JSON.
5. Buka halaman **Upload Runner**, upload mapping JSON, lalu upload Excel.

## Cara live test upload ke Oracle Fusion

Gunakan instance DEV/TEST dulu.

1. Buka halaman **Upload Runner**.
2. Matikan **Mock / Dry Run mode**.
3. Isi:
   - Oracle Base URL, contoh: `https://your-instance.oraclecloud.com`
   - Username
   - Password
4. Klik **Test Connection**.
5. Klik **Test Selected Row** untuk 1 baris dulu.
6. Kalau berhasil, baru klik **Run All Rows**.

## Field Name vs Field ID

Untuk POST/create, app sekarang memakai field ID sebagai default:

```text
ManagementBusinessUnitId
LegalEntityId
ProfitCenterBusinessUnitId
ItemDefinitionOrganizationId
```

Field berikut diperlakukan sebagai read-only untuk POST dan akan divalidasi error kalau dipakai:

```text
ManagementBusinessUnitName
LegalEntityName
ProfitCenterBusinessUnitName
ItemDefinitionOrganizationName
```

## Secrets lokal opsional

Copy contoh secrets:

```bash
cp .streamlit/secrets.example.toml .streamlit/secrets.toml
```

Lalu isi:

```toml
ORACLE_BASE_URL = "https://your-instance.oraclecloud.com"
ORACLE_USERNAME = "your.username"
ORACLE_PASSWORD = "your-password"
```

Jangan commit `.streamlit/secrets.toml`.

## Catatan schema

File schema utama ada di:

```text
schemas/inventory_organizations.json
```

Kamu bisa menambah field atau membuat schema endpoint lain dengan pola yang sama.

Catatan penting: `OrganizationId` dibuat oleh Oracle saat organization dibuat. Di aplikasi ini, `OrganizationId` diperlakukan sebagai field hasil response/log, bukan input Excel.

## File penting

```text
app.py
pages/1_Template_JSON_Builder.py
pages/2_Upload_Runner.py
pages/3_Reference_From_Fusion.py
schemas/inventory_organizations.json
services/config_loader.py
services/excel_service.py
services/payload_builder.py
services/validation_service.py
services/oracle_client.py
services/reference_service.py
sample_data/inventory_organizations_sample.csv
sample_data/inventory_organizations_mapping.json
sample_data/inventory_organizations_sample_payload.json
```

## Safety checklist sebelum production

- Pakai DEV/TEST dulu.
- Tambahkan role/access control jika dipakai banyak user.
- Simpan credential di secret manager, bukan di file biasa.
- Aktifkan audit log permanen jika masuk production.
- Batasi akses PROD dan gunakan approval sebelum upload massal.
