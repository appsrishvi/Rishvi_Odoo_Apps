# -*- coding: utf-8 -*-
# © 2025 Rishvi LTD – MIT/AGPL‑3

from odoo import models, fields, _
from markupsafe import escape
from collections import defaultdict


class TimesheetDailyMail(models.Model):
    _name = "timesheet.daily.mail"
    _description = "Daily Timesheet Email Sender"
    _inherit = ["mail.thread"]

    def send_daily_email(self):
        today = fields.Date.context_today(self)
        timesheets = self.env["account.analytic.line"].search(
            [("date", "=", today)]
        )
        if not timesheets:
            return

        employee_map = defaultdict(list)
        total_map = {}
        for ts in timesheets:
            emp = ts.employee_id
            if not emp:
                continue
            emp_name = emp.name
            employee_map[emp_name].append({
                "project": ts.project_id.name or "",
                "task": ts.task_id.name or "",
                "description": ts.name or "",
                "hours": ts.unit_amount or 0.0,
            })

        # ----------------- Summary Table ------------------
        summary_rows = []
        for employee, entries in employee_map.items():
            total = sum(e["hours"] for e in entries)
            total_map[employee] = total
            summary_rows.append(f"""
                <tr>
                    <td style="border:1px solid #ccc;padding:8px;">{escape(employee)}</td>
                    <td style="border:1px solid #ccc;padding:8px;text-align:right;font-weight:bold;color:#2c3e50;">
                        {total:.2f}
                    </td>
                </tr>
            """)

        summary_table = f"""
            <h3 style="font-family:Arial,sans-serif;margin-top:0;">
                {_("Summary of Employee Hours")}
            </h3>
            <table style="width:100%;border-collapse:collapse;font-family:Arial,sans-serif;font-size:14px;">
                <thead>
                    <tr style="background:#f4f4f4;">
                        <th style="border:1px solid #ccc;padding:8px;text-align:left;">{_('Employee')}</th>
                        <th style="border:1px solid #ccc;padding:8px;text-align:right;">{_('Total Hours')}</th>
                    </tr>
                </thead>
                <tbody>
                    {''.join(summary_rows)}
                </tbody>
            </table>
        """

        # ----------------- Detail Tables ------------------
        details_html = []
        for employee, entries in employee_map.items():
            rows = "".join(f"""
                <tr>
                    <td style="border:1px solid #ccc;padding:8px;word-wrap:break-word;max-width:200px;">{escape(e["project"])}</td>
                    <td style="border:1px solid #ccc;padding:8px;">{escape(e["task"])}</td>
                    <td style="border:1px solid #ccc;padding:8px;">{escape(e["description"])}</td>
                    <td style="border:1px solid #ccc;padding:8px;text-align:right;">{e["hours"]:.2f}</td>
                </tr>
            """ for e in entries)

            details_html.append(f"""
                <h3 style="font-family:Arial,sans-serif;color:#2c3e50;margin:18px 0 6px;">
                    &nbsp;{escape(employee)}
                </h3>
                <table style="width:100%;border-collapse:collapse;font-family:Arial,sans-serif;font-size:13px;">
                    <thead>
                        <tr style="background:#f2f2f2;">
                            <th style="border:1px solid #ccc;padding:8px;text-align:left;">Project</th>
                            <th style="border:1px solid #ccc;padding:8px;text-align:left;">Task</th>
                            <th style="border:1px solid #ccc;padding:8px;text-align:left;">Description</th>
                            <th style="border:1px solid #ccc;padding:8px;text-align:right;">Hours</th>
                        </tr>
                    </thead>
                    <tbody>{rows}</tbody>
                </table>
            """)

        # ----------------- Final Email Body ------------------
        body_html = f"""
            <p style="font-family:Arial,sans-serif;font-size:14px;color:#333;">
                {_("Hello")},<br/><br/>
                {_("Here is the")} <strong>{_("Daily Timesheet Report")}</strong>
                {_("for")} <strong>{today.strftime('%B %d, %Y')}</strong>.
            </p>
            {summary_table}
            <hr style="margin:25px 0;" />
            {''.join(details_html)}
            <p style="font-family:Arial,sans-serif;font-size:14px;color:#333;margin-top:20px;">
                {_("Regards")},<br/>
                <strong>Odoo Timesheet System</strong>
            </p>
        """

        if self.env.user.id != self.env.ref('base.user_root').id:
            # real user
            sender_email = self.env.user.company_id.partner_id.email or self.env.user.email
        else:
            # fallback to admin user
            admin_user = self.env.ref("base.user_admin", raise_if_not_found=False)
            sender_email = admin_user.company_id.partner_id.email or admin_user.email
        mail_values = {
            "subject": _("Daily Timesheet Report – %s") % today.strftime("%d %b %Y"),
            "body_html": body_html,
            "email_to": sender_email,
            "email_from": sender_email,
        }
        self.env['mail.mail'].create(mail_values).send()
