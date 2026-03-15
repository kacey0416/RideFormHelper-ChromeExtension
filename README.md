# SafeRide Form Helper

Manifest V3 Chrome extension for faster Safe Ride Smartsheet submissions.


## Current behavior

- Stores partner profile locally (`chrome.storage.local`):
  - Partner Name
  - Partner Number (digits only)
  - Email (valid email format)
  - Store Number (digits only)
  - Province
  - Electronic Signature
- Profile state badge:
  - `OK` when all required inputs are valid
  - `!` when any required input is missing/invalid
- Accepts receipt image uploads with limits:
  - Max `30` files in queue
  - Max `4.0 MB` per file
- Runs OCR sequentially for newly uploaded files.
- Parses and stores:
  - Date of ride
  - Time of ride
  - Cost
- Queue supports status states:
  - `pending`, `parsing`, `ready`, `filled`, `submitted`, `error`
- Selecting a queue item auto-fills the Smartsheet form when parsed values are available.
- `Show` preview button appears only for the selected item when status is complete (`ready`/`filled`/`submitted`).
- Detects successful Smartsheet submit (Thank-you message), marks current item as `submitted`, and shows `Back to Form` button.
- Profile + queue data uses TTL cleanup (auto-expire after 3 days).

## Important limitations

- OCR accuracy is not guaranteed. Always verify before final submit.
- Smartsheet DOM/field structure changes can break selectors.
- Submission itself is still user-driven (user clicks Smartsheet submit).
- OCR language data is fetched from `tessdata.projectnaptha.com`.


## Usage

1. Open the Safe Ride form page.  
2. Enter Partner Info and click `Save`; the profile is saved to local storage.  
3. Upload receipt screenshots (up to 30 files).  
4. When parsing finishes, the first receipt is automatically filled into the form.  
5. Before submitting, use `Show` preview to confirm the receipt details are correct, then submit.  
6. After submission is confirmed, that receipt is automatically marked as completed in the list.  
7. Return to the form page, and the next receipt is auto-filled.  
8. Repeat until all receipts are submitted.

