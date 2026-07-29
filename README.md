# Order Haravan Report

Static GitHub Pages report generated from local `Orders_T*.xlsx` files.

## Update Workflow

1. Update the monthly Excel files in this folder.
2. Run `generate_order_report.py` to rebuild `order_report_from_raw_data.html`.
3. Commit `generate_order_report.py` and `order_report_from_raw_data.html` if changed.
4. Push to `main`; GitHub Pages will update the public report link.

## Important Size Rule

Do not embed product images as base64/data URIs inside the HTML report.

Images must stay as separate files under `product_images/`, and report records should reference paths such as `product_images/SKU.png`. Embedding images into `REPORT_DATA.meta.imageAssets` makes `order_report_from_raw_data.html` exceed GitHub's 100 MB file limit.

The generator should keep:

```python
"imageAssets": {},
```

and should not call a base64 image asset builder.

Report URL:

```text
https://scottvinguyen.github.io/Order-Haravan-Report/
```
