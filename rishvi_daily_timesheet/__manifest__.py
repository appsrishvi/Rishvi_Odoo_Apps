{
    "name": "Daily Timesheet Report Email",
    "summary": "Send daily timesheet reports via email",
    'author': "Rishvi Ltd",
    'website': "https://rishvi.co.uk/",
    'version': '18.0.0.1',
    "category": "Timesheets",
    "depends": ["hr_timesheet", "mail"],
    "data": [
        'security/ir.model.access.csv',
        "data/timesheet_report_cron.xml",
        "wizard/daily_timsheet_wizard.xml",
    ],
    'images': ['static/description/banner.png'],
    'license': 'AGPL-3',
    'installable': True,
    'auto_install': False,
    'application': True,
}
