## Assumptions

- 100 staff members use the app
- 100 incoming appointment requests and appointment assignments per hour
- 10 appointment reassignments per hour
- 100 appointment details views per hour (1 per appointment request)
- 100 tab views per hour (1 per staff member per hour)
- Staff member operations are spread out evenly over an 8 hour period each day
- Operations associated with app installation are negligible, since they only occur once per user/team

## SKU Recommendations

- App Service: Standard (S1)

## Estimated Usage

### Cosmos DB throughput

The table below lists the main operations of the app and their corresponding Cosmos operations. Infrequent operations are ignored since the number of reads/writes/updates is negligible.

| App operation | # reads | # writes | # updates | operations/hr |
| ------------- | ------- | -------- | --------  | ------------- |
| Create appointment request | 2 | 1 | | 100 |
| Assign appointment request | 4 | | 1 | 100 |
| Complete appointment request | 1 | | 1 | 100 |
| View appointment details | 1 | | | 100 |
| View tab | 1 | | | 100 |
| Total/hr | 900 | 100 | 200 | |

Even at 5 RU for a Cosmos operation (the expected cost of a 1 KB write), this comes out to < 2 RU/s.

### Cosmos DB storage
- Documents generally won't exceed 5 KB
- 100 appointments/hr * 8 hrs/day * 365 days/year * 5 KB/appointment = 1.5 GB/year

## Estimated cost

**IMPORTANT:** This is only an estimate, based on the assumptions above. Your actual costs may vary.

Prices were taken from the [Azure Pricing Overview](https://azure.microsoft.com/en-us/pricing/) on 2 November 2020, for the West US 2 region.

Use the [Azure Pricing Calculator](https://azure.com/e/c3bb51eeb3284a399ac2e9034883fcfa) to model different service tiers and usage patterns.

| Resource   | Tier     | Usage / Month                | Price / Month (USD)      |
|------------|----------|------------------------------|--------------------------|
| App Service Plan | S1 | 744 hours (31 days) | $74.40 |
| App Service (bot, API, etc.) | - | - | (charged to App Service Plan) |
| Cosmos DB | - | 2 GB, 400 RU/s | $24.31 |
| Bot Channels Registration | F0 | N/A | Free (for Teams channel) |
| Azure Monitor / Application Insights | - | < 5GB data | Free (up to 5 GB)
| **Total**  | | | **$98.71** |
