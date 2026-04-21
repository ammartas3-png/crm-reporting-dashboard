# crm-reporting-dashboard

CRM Reporting Tool for FTD, CR, performance analytics, agent calls, check-in/check-out and country AFF report.

## Yeni: CRM + PowerBI rapor otomasyonu

Bu repoya `report_automation.py` eklendi. Script temel olarak iki dosya ile çalışır, opsiyonel olarak 3. bir dosya da alabilir:

- CRM export (csv/xlsx)
- PowerBI export (csv/xlsx)
- Purple source export (csv/xlsx, opsiyonel - örn. Comments / Call Att gibi mor sütunlar)

### Ne yapar?

1. CRM dosyasındaki müşteri numarası boş olan satırları PowerBI verisiyle doldurmaya çalışır.
2. Yorum eşleştirmesinde:
   - önce exact eşleşme,
   - yoksa fuzzy (yaklaşık) eşleşme kullanır.
3. Mor sütunları ayrı dosyadan (`--purple`) CRM'e join ederek getirir (opsiyonel).
4. Lead bazında kaç kere arandığını hesaplar ve `lead_call_count` kolonu ekler.
5. Aşağıdaki özet tabloları üretir:
   - `lead_call_counts`
   - `call_count_distribution`
   - `aff_status_ratios`
   - `call_frequency_by_aff_status`
6. Çıktıyı tek bir Excel dosyasına (çoklu sheet) yazar.

### Kurulum

```bash
python -m pip install -r requirements.txt
```

### Konfigürasyon

`config.example.yml` dosyasını kopyalayıp kendi kolon isimlerine göre güncelle:

```bash
cp config.example.yml config.yml
```

### Çalıştırma

```bash
python report_automation.py \
  --crm ./data/crm.xlsx \
  --powerbi ./data/powerbi.xlsx \
  --purple ./data/purple_source.xlsx \
  --config ./config.yml \
  --output ./report_output.xlsx
```

`--purple` parametresi opsiyoneldir. Verirsen, config içindeki `purple_source` ayarına göre join yapılır.

`purple_source.stage` seçenekleri:
- `before_customer_match`: purple join önce yapılır
- `after_customer_match`: önce customer no bulunur, sonra purple join yapılır (çoğu durumda önerilen)

### Çıktı sayfaları

- `crm_enriched`: CRM satırları + doldurulan müşteri numarası + eşleşme bilgileri
- `match_audit`: hangi satır nasıl eşleşti (exact/fuzzy/not_found)
- `purple_merge_audit`: mor sütun merge işlemi özeti (kaç satır eşleşti/uygulandı)
- `lead_call_counts`: lead başına arama sayısı
- `call_count_distribution`: kaç lead kaç kez aranmış dağılımı
- `aff_status_ratios`: AFF + status bazında adet ve oranlar
- `overall_status_summary`: tüm status dağılımı (adet + oran)
- `call_frequency_by_aff_status`: AFF/status/call_count kırılımında lead sayısı
- `dashboard_summary`: görseldeki gibi formatlanmış toplu özet tablolar
