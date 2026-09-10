/**
 * Everything on the invoice that comes from Orion Corp rather than from the
 * uploaded spreadsheet. The legacy Django app hard-coded these strings inside
 * `generate_pdf.py`; they live here so they can be corrected in one place.
 *
 * The legacy source also built a Punjab & Sind Bank block and then discarded it
 * without rendering — the ICICI block below is the one that actually appeared on
 * issued invoices, and is the one confirmed as current.
 */
export const COMPANY = {
  name: "Orion Corp",
  address:
    "P-2/03, Tower 3B, Purvanchal Silver City -2, Sector Pi-2, Greater Noida, India, 201308, Phone:0120-4543418",
  warehouse:
    "WAREHOUSE- PLOT NO-239, Giani Compound, Giani Boarder, Opposite Metro Pillar No.160, Behind Giani Gill Transport, Post Ckikamberpur",
  gstin: "09AKGPG4906P1ZQ",
  state: "State Name: Uttar Pradesh, Code: 09",
  bank: {
    accountName: "Orion Corp",
    line: "ICICI BANK LTD A/C no.: 003105037390",
    ifsc: "IFSC Code: ICIC0000031",
    branchAddress:
      "K-1,SENIOR MALL, SECTOR 18, NOIDA, UTTAR PRADESH, PIN CODE : 201301",
  },
  msmeNumber: "UDYAM-UP-29-0008397",
  gemSellerId: "C0FB200001154517",
  /**
   * The legacy PDF printed "for Orion the Corp". Corrected to the real company
   * name for invoices issued from this app onwards.
   */
  signatureFor: "for Orion Corp",
  declaration:
    "We declare that this invoice shows the actual price of the goods described and that all particulars are true and correct.",
  footer: "This is a Computer Generated Invoice",
} as const;
