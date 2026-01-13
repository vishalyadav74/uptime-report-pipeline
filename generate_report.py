<td class="kpi"
    title="Weekly Uptime Avg of all accounts">
  <div class="kpi-value green">✔ {{ overall_uptime }}</div>
  Overall Uptime Avg
</td>

<td class="kpi"
    title="Total outages in weekly">
  <div class="kpi-value red">⚠ {{ outage_count }}</div>
  Outages
</td>

<td class="kpi"
    title="{{ major_incident.hover }}">
  <div class="kpi-value purple">🚩 {{ major_incident.account }}</div>
  Most Affected Account
</td>

<td class="kpi"
    title="Outage downtime for the most affected account">
  <div class="kpi-value orange">⏱ {{ total_downtime }} mins</div>
  Outage Downtime
</td>
