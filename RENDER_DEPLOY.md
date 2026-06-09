# OpsPilot AI - Render deploy

## Render settings

Build Command:
```bash
pip install -r requirements.txt
```

Start Command:
```bash
gunicorn app:app
```

Environment variables:
```text
PYTHON_VERSION=3.11.9
```

Optional Stripe variables:
```text
STRIPE_PUBLISHABLE_KEY=
STRIPE_SECRET_KEY=
STRIPE_WEBHOOK_SECRET=
STRIPE_PRO_PRICE_ID=
APP_BASE_URL=https://your-render-url.onrender.com
```

## Local testing

```bash
python app.py
```

PC:
```text
http://127.0.0.1:5000
```

Phone on same Wi-Fi:
```text
http://YOUR-PC-IP:5000
```

Admin:
```text
admin@opspilot.ai
admin123
```

## Included final features

- Dashboard
- Morning Brief Builder
- Mileage
- Expenses
- Yard Check
- Yard Check search/edit/delete
- Yard Check batch add
- Bulk voice trailer logging
- Take Photo / Upload Photo
- Handover with photo/camera
- Team with custom roles
- Plans Free/Pro
- Stripe Pro subscription setup
- Admin Billing
- Admin manual Pro
- Admin 30 day Pro trial
- Admin Gift Pro
- Pro-only Yard Configuration
- Free fixed Door 1-100 / Fence 1-120
- PWA app install support
