# JSON Data Files Reference
The following attributes are accepted via the json data file.

- case_name:
    - Optional - str
    - Name of the test case
    - default: `f"test_{datetime.datetime.now().strftime('%m%d%Y_%H%M%S')}"`
- description:
    - Optional - str
    - Detailed description of the test case 
    - default: Empty string - `""`
- business_owner:
    - Optional - str
    - Name of the responsible Business Process Owner or Key User 
    - default: `"Business Process Owner"`
- it_owner:
    - Optional - str 
    - Name of the responsible IT user 
    - default: `"Technical Owner"`
- doc_link:
    - Optional - str
    - URL link to detailed documentation of the test case
    - default: Empty string - `""`
- case_path:
    - Optional - str
    - Path to the root of the test case directory
    - default: Uses the directory of the current test case script file
- date_format:
    - Optional - str 
    - Date string format used in SAP GUI date fields
    - default: `"%m/%d/%Y"`
- explicit_wait:
    - Optional - float
    - Amount of time in seconds to wait when a function has a @explicit_wait_before or @explicit_wait_after decorator
    - default: `0.25`
- screenshot_on_pass:
    - Optional - bool
    - Flag controlling the capture of screenshots when a step is marked passing
    - default: `False`
- screenshot_on_fail:
    - Optional - bool
    - Flag controlling the capture of screenshots when a step is marked failing
    - default: `False`
- fail_on_error:
    - Optional - bool
    - Flag controlling how an unexpected technical python error occurring during a step is handled
        - True the error is captured as an error in the logs and the step is marked failing 
        - False the error is captured as a warning in the logs and the step is marked passing
    - default: `True`
- exit_on_fail:
    - Optional - bool
    - Flag controlling how a script responds to a failing test step
        - True the error is captured as an error in the logs and the script exits
        - False the error is captured as a warning in the logs and the script continues 
    - default: `True`
- close_on_cleanup:
    - Optional - bool
    - Flag controlling how the SAP GUI is handled when the atexit function executes
        - The atexit function is run at the ending of the script independent of how the script is terminated
        - True the SAP GUI will be closed if at all possible
        - False the SAP GUI will remain open
    - default: `True`
- system: 
    - Optional - dict or str
    - System used when opening a connection to the SAP GUI Scripting Engine API
    - default: `""`
        - The following key:value pairs are accepted as part of a dictionary supplying multiple systems:
            - "erp": "string of erp system"
            - "ewm": "string of ewm system"
            - "hcm": "string of hcm system"
            - "apo": "string of apo system"
            - "gts": "string of gts system"
            - "solman": "string of solman system"
            - "gateway": "string of gateway system"
- data:
    - Optional - dict or None
    - Dictionary of key:value pairs providing detailed test case data
    - default: `None`

## Data Examples
The following are examples of data sections for various cases.

### Sales Inquiry
```json

```

### Sales Quote
```json

```

### Sales Order
```json
{
    "order_type": "",
    "sales_org": "",
    "dist_ch": "",
    "division": "",
    "sales_office": "",
    "sales_group": "",
    "sold_to": "",
    "ship_to": "",
    "customer_ref": "",
    "customer_ref_date": "",
    "requested_delivery_date": "",
    "complete_delivery": "",
    "delivery_block": "",
    "billing_block": "",
    "pricing_date": "",
    "payment_terms": "",
    "inco_version": "",
    "incoterms": "",
    "inco1": "",
    "inco2": "",
    "order_reason": "",
    "plant": "",
    "version": "",
    "guarantee": "",
    "delivery_time": "",
    "doc_currency": "",
    "customer_group": "",
    "price_list_type": "",
    "usage": "",
    "customer_price_group": "",
    "sales_district": "",
    "unloading_point": "",
    "receiving_point": "",
    "department": "",
    "order_combination": "",
    "shipping_type": "",
    "special_process_indicator": "",
    "shipping_condition": "",
    "means_of_transport_type": "",
    "means_of_transport": "",
    "pod_relevant": "",
    "fixed_value_date": "",
    "additional_value_days": "",
    "invoicing_dates": "",
    "manual_invoice_maintenance": "",
    "service_rendered_date": "",
    "tax_departure_country": "",
    "alternative_tax_classification": "",
    "tax_departure_country": "",
    "tax_destination_country": "",
    "triangular_deal_within_eu_indicator": "",
    "items": [
        {
            "material": "",
            "qty": "", 
            "uom": "",
            "item_category": "",
            "storage_location": "",
            "amount": "", 
            "shipping_point": "",
            "pricing_conditions": [
                {
                    "pricing_condition": "",
                    "pricing_amount": ""
                },
                {
                    "pricing_condition": "",
                    "pricing_amount": ""
                }
            ]
        },
        {
            "material": "",
            "qty": "", 
            "uom": "",
            "item_category": "",
            "storage_location": "",
            "amount": "", 
            "shipping_point": "",
            "pricing_conditions": [
                {
                    "pricing_condition": "",
                    "pricing_amount": ""
                },
                {
                    "pricing_condition": "",
                    "pricing_amount": ""
                }
            ]
        }
    ]
}
```

### Return Order
```json

```

### Outbound Delivery
```json

```

### Inbound Delivery
```json

```

### Return Delivery
```json

```

### Purchase Req
```json

```

### Purchase Order
```json
{
	"po_type":"NB",
	"vendor": "",
	"purchase_org": "",
	"purchasing_group": "",
	"company_code": "",
	"incoterms": "",
	"incoterms_location_1": "",
	"items": [
		{
			"material": "",
			"qty": "",
			"uom": "",
			"net_price": "",
			"currency": "",
			"plant": "",
			"storage_location": "",
			"price_unit": "",
			"opu": ""
		},
		{
			"material": "",
			"qty": "",
			"uom": "",
			"net_price": "",
			"currency": "",
			"plant": "",
			"storage_location": "",
			"price_unit": "",
			"opu": ""
		}
	],
	"po": "",
	"inbound_deliveries": [
		{
			"delivery": "",
			"idoc": ""
		},
		{
			"delivery": "",
			"idoc": ""
		},
		{
			"delivery": "",
			"idoc": ""
		}
	]
}
```

### Invoice
```json

```

### Production Order
```json

```

### Material Master
```json

```

### Bill of Material
```json

```

### Info Record
```json

```
