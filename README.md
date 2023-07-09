# SAP GUI Framework

A Framework library for controlling the SAP GUI Desktop and focused on testing business processes. The library uses the native SAP GUI scripting engine for interaction with the desktop client application. 

Other dependances include the below packages as this library would not be possible without the hard work of so many others. 

pywin32 by Mark Hammond

What makes this library different from other SAP GUI libraries? 

The main difference is the focus on testing end to end business processes in and around SAP GUI. While other libraries are similar in many aspects there are many additional comfort of life functions included that aren't in other libraries. 

This package is available also on PyPi and available for install via pipenv. For the latest updates please use the git install process below. 

If you haven't already created a virtual environment with pipenv you do the following steps first:
1. Create virtual environment: `pipenv --python 3.11`
2. Enter new virtual environment: `pipenv shell`

## To install the SapGuiFramework
```shell
pipenv install 'SapGuiFramework @ git+https://github.com/jduncan8142/SapGuiFramework.git@main'
```

## To update the SapGuiFramework
```shell
pipenv uninstall sapguiframework; pipenv install 'SapGuiFramework @ git+https://github.com/jduncan8142/SapGuiFramework.git@main'
```
## Note
If you have not already you likely will need Scripting Tracker from Stefan Schnell. You can find it at https://tracker.stschnell.de/

## Keywords Documentation
### Data Files:
The following attributes are accepted via the json data file for the test case: 
- case_name {str} -- Name of the test case (default: {f"test_{datetime.datetime.now().strftime('%m%d%Y_%H%M%S')}"})
- description {str} -- Detailed description of the test case (default: {""})
- business_owner {str} -- Name of the Business Process Owner or Key User (default: {"Business Process Owner"})
- it_owner {str} -- Name of the IT responsible (default: {"Technical Owner"})
- doc_link {str} -- URL link to detailed documentation of the test case (default: {""})
- case_path {str} -- Path to the root of the test case directory (default: {""})
- date_format {str} -- _description_ (default: {"%m/%d/%Y"})
- explicit_wait {float} -- _description_ (default: {0.25})
- screenshot_on_pass {bool} -- _description_ (default: {False})
- screenshot_on_fail {bool} -- _description_ (default: {False})
- fail_on_error {bool} -- _description_ (default: {True})
- exit_on_fail {bool} -- _description_ (default: {True})
- close_on_cleanup {bool} -- _description_ (default: {True})
- system {dict|str} -- System to using when opening a connection to the SAP GUI API (default: {""})
    - The following key:value pairs are accepted as part of a dictionary to supply multiple systems:
        - erp {str}
        - ewm {str}
        - hcm {str}
        - apo {str}
        - gts {str}
        - solman {str}
        - gateway {str}
- Data {dict|None} -- Dictionary of key:value pairs providing detailed test case data (default: {None})
    - Sales Inquiry:
    - Sales Quote:
    - Sales Order:
        - order_type {str}
        - sales_org {str}
        - dist_ch {str}
        - division {str}
        - sales_office {str}
        - sales_group {str}
        - sold_to {str}
        - ship_to {str}
        - customer_ref {str}
        - customer_ref_date {str}
        - requested_delivery_date {str}
        - complete_delivery {bool}
        - delivery_block {str}
        - billing_block {str}
        - pricing_date {str}
        - payment_terms {str}
        - inco_version {str}
        - incoterms {str}
        - inco1 {str}
        - inco2 {str}
        - order_reason {str}
        - plant {str}
        - version {str}
        - guarantee {str}
        - delivery_time {str}
        - doc_currency {str}
        - customer_group {str}
        - price_list_type {str}
        - usage {str}
        - customer_price_group {str}
        - sales_district {str}
        - unloading_point {str}
        - receiving_point {str}
        - department {str}
        - order_combination {bool}
        - shipping_type {str}
        - special_process_indicator {str}
        - shipping_condition {str}
        - means_of_transport_type {str}
        - means_of_transport {str}
        - pod_relevant {bool}
        - fixed_value_date {str}
        - additional_value_days {str}
        - invoicing_dates {str}
        - manual_invoice_maintenance {bool}
        - service_rendered_date {str}
        - tax_departure_country {str}
        - alternative_tax_classification {str}
        - tax_departure_country {str}
        - tax_destination_country {str}
        - triangular_deal_within_eu_indicator {bool}
        - items {list}
            - Items is a list of dictionaries with the following keys supported:
                - material {str}
                - qty {str}
                - uom {str}
                - item_category {str}
                - storage_location {str}
                - amount {str}
                - shipping_point {str}
                - pricing_conditions {list}
                    - pricing_condition {str}
                    - pricing_amount {str}
    - Return Order
    - Outbound Delivery:
    - Inbound Delivery:
    - Return Delivery:
    - Purchase Req:
    - Purchase Order:
    - Invoice:
    - Production Order:
    - Material Master:
    - Bill of Material:
    - Info Record:

### .env Files
Coming Soon!

### Tests
Coming Soon!

### PyPi Build Process
Coming Soon!