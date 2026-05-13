name: Freshdesk Full Export

on:
  workflow_dispatch:  # Solo manual, no tiene schedule

jobs:
  export:
    runs-on: ubuntu-latest
    steps:
      - uses: actions/checkout@v4
      - uses: actions/setup-python@v5
        with:
          python-version: "3.12"
      - name: Instalar dependencias
        run: pip install requests pandas openpyxl
      - name: Ejecutar exportación completa
        env:
          FRESHDESK_DOMAIN:  ${{ secrets.FRESHDESK_DOMAIN }}
          FRESHDESK_API_KEY: ${{ secrets.FRESHDESK_API_KEY }}
          MAIL_FROM:         ${{ secrets.MAIL_FROM }}
          MAIL_PASSWORD:     ${{ secrets.MAIL_PASSWORD }}
          MAIL_TO:           ${{ secrets.MAIL_TO }}
        run: python freshdesk_export_full.py
      - uses: actions/upload-artifact@v4
        with:
          name: freshdesk-completo-${{ github.run_id }}
          path: freshdesk_todos_tickets_*.xlsx
          retention-days: 30
    main()
