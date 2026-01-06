import streamlit as st
import pandas as pd
from pathlib import Path
import sys
from io import BytesIO

# Add parsers and core to path
sys.path.append(str(Path(__file__).parent))

from parsers.reconcile_excel import parse_reconciliation_excel
from parsers.invoice_xml import parse_invoice_xml
from parsers.invoice_excel import parse_invoice_excel
from parsers.invoice_pdf import parse_invoice_pdf
from parsers.invoice_ocr import parse_invoice_image
from core.normalize import normalize_data
from core.compare import compare_datasets
from core.report import generate_excel_report

st.set_page_config(
    page_title="Invoice Reconciliation Tool",
    page_icon="📊",
    layout="wide"
)

st.title("📊 Invoice Reconciliation Tool")
st.markdown("Compare reconciliation Excel files with invoices (XML, Excel, PDF, or Image)")

# Sidebar options
st.sidebar.header("⚙️ Comparison Settings")
vat_tolerance = st.sidebar.number_input(
    "VAT Tolerance (%)",
    min_value=0.0,
    max_value=10.0,
    value=0.0,
    step=0.1,
    help="Allowed percentage difference for VAT amounts"
)
amount_tolerance = st.sidebar.number_input(
    "Amount Tolerance (%)",
    min_value=0.0,
    max_value=10.0,
    value=0.0,
    step=0.1,
    help="Allowed percentage difference for line amounts"
)
fuzzy_match = st.sidebar.checkbox(
    "Enable Fuzzy Product Name Matching",
    value=False,
    help="Use RapidFuzz for flexible product name matching"
)
compare_after_discount = st.sidebar.checkbox(
    "Compare After Discount",
    value=False,
    help="Include discount columns in comparison"
)

# File upload section
col1, col2 = st.columns(2)

with col1:
    st.subheader("1️⃣ Upload Reconciliation Excel")
    reconciliation_file = st.file_uploader(
        "Choose reconciliation Excel file",
        type=['xlsx', 'xls'],
        key="reconciliation"
    )

with col2:
    st.subheader("2️⃣ Upload Invoice File")
    invoice_file = st.file_uploader(
        "Choose invoice file",
        type=['xml', 'xlsx', 'xls', 'pdf', 'png', 'jpg', 'jpeg'],
        key="invoice"
    )

if reconciliation_file and invoice_file:
    try:
        with st.spinner("🔄 Processing files..."):
            
            # Parse reconciliation file
            st.info("📖 Parsing reconciliation Excel...")
            reconciliation_data = parse_reconciliation_excel(reconciliation_file)
            
            # Parse invoice file based on extension
            invoice_ext = Path(invoice_file.name).suffix.lower()
            st.info(f"📖 Parsing invoice file ({invoice_ext})...")
            
            if invoice_ext == '.xml':
                invoice_data = parse_invoice_xml(invoice_file)
            elif invoice_ext in ['.xlsx', '.xls']:
                invoice_data = parse_invoice_excel(invoice_file)
            elif invoice_ext == '.pdf':
                invoice_data = parse_invoice_pdf(invoice_file)
            elif invoice_ext in ['.png', '.jpg', '.jpeg']:
                invoice_data = parse_invoice_image(invoice_file)
            else:
                st.error(f"Unsupported file format: {invoice_ext}")
                st.stop()
            
            # Normalize data
            st.info("🔧 Normalizing data...")
            normalized_reconciliation = normalize_data(
                reconciliation_data['line_items'],
                compare_after_discount=compare_after_discount
            )
            normalized_invoice = normalize_data(
                invoice_data['line_items'],
                compare_after_discount=compare_after_discount
            )
            
            # Compare datasets
            st.info("🔍 Comparing data...")
            comparison_result = compare_datasets(
                normalized_reconciliation,
                normalized_invoice,
                reconciliation_data.get('totals', {}),
                invoice_data.get('totals', {}),
                vat_tolerance=vat_tolerance / 100,
                amount_tolerance=amount_tolerance / 100,
                fuzzy_match=fuzzy_match
            )
        
        st.success("✅ Processing complete!")
        
        # Display extracted data side by side
        st.header("📋 Extracted Data")
        col1, col2 = st.columns(2)
        
        with col1:
            st.subheader("Reconciliation Data")
            st.dataframe(
                reconciliation_data['line_items'],
                use_container_width=True,
                height=300
            )
            if reconciliation_data.get('totals'):
                st.json(reconciliation_data['totals'])
        
        with col2:
            st.subheader("Invoice Data")
            st.dataframe(
                invoice_data['line_items'],
                use_container_width=True,
                height=300
            )
            if invoice_data.get('totals'):
                st.json(invoice_data['totals'])
        
        # Display comparison results
        st.header("🔍 Comparison Results")
        
        # Summary metrics
        col1, col2, col3, col4 = st.columns(4)
        
        with col1:
            total_items = len(comparison_result['comparison_table'])
            st.metric("Total Items", total_items)
        
        with col2:
            matched = comparison_result['summary']['matched_items']
            st.metric("Matched Items", matched, delta=f"{matched/total_items*100:.1f}%" if total_items > 0 else "0%")
        
        with col3:
            mismatched = comparison_result['summary']['mismatched_items']
            st.metric("Mismatched Items", mismatched, delta=f"-{mismatched/total_items*100:.1f}%" if total_items > 0 else "0%")
        
        with col4:
            totals_match = comparison_result['summary']['totals_match']
            st.metric("Totals Match", "✅ YES" if totals_match else "❌ NO")
        
        # Comparison table with styling
        st.subheader("📊 Detailed Comparison")
        
        comparison_df = comparison_result['comparison_table']
        
        def highlight_mismatches(row):
            """Highlight mismatched rows in red, matched in green"""
            if row['status'] == 'MATCH':
                return ['background-color: #d4edda'] * len(row)
            elif row['status'] in ['MISMATCH', 'MISSING_IN_INVOICE', 'EXTRA_IN_INVOICE']:
                return ['background-color: #f8d7da'] * len(row)
            else:
                return [''] * len(row)
        
        styled_df = comparison_df.style.apply(highlight_mismatches, axis=1)
        st.dataframe(styled_df, use_container_width=True, height=400)
        
        # Totals comparison
        st.subheader("💰 Totals Comparison")
        totals_df = pd.DataFrame([
            {
                'Item': 'Total Before Tax',
                'Reconciliation': comparison_result['totals_comparison'].get('reconciliation_total_before_tax', 'N/A'),
                'Invoice': comparison_result['totals_comparison'].get('invoice_total_before_tax', 'N/A'),
                'Match': '✅' if comparison_result['totals_comparison'].get('total_before_tax_match') else '❌'
            },
            {
                'Item': 'VAT Rate (%)',
                'Reconciliation': comparison_result['totals_comparison'].get('reconciliation_vat_rate', 'N/A'),
                'Invoice': comparison_result['totals_comparison'].get('invoice_vat_rate', 'N/A'),
                'Match': '✅' if comparison_result['totals_comparison'].get('vat_rate_match') else '❌'
            },
            {
                'Item': 'VAT Amount',
                'Reconciliation': comparison_result['totals_comparison'].get('reconciliation_vat_amount', 'N/A'),
                'Invoice': comparison_result['totals_comparison'].get('invoice_vat_amount', 'N/A'),
                'Match': '✅' if comparison_result['totals_comparison'].get('vat_amount_match') else '❌'
            },
            {
                'Item': 'Total Payment',
                'Reconciliation': comparison_result['totals_comparison'].get('reconciliation_total_payment', 'N/A'),
                'Invoice': comparison_result['totals_comparison'].get('invoice_total_payment', 'N/A'),
                'Match': '✅' if comparison_result['totals_comparison'].get('total_payment_match') else '❌'
            }
        ])
        st.dataframe(totals_df, use_container_width=True)
        
        # Export report
        st.header("📥 Export Report")
        
        if st.button("📄 Generate Excel Report", type="primary"):
            with st.spinner("Generating report..."):
                report_buffer = generate_excel_report(
                    reconciliation_data,
                    invoice_data,
                    comparison_result
                )
                
                st.download_button(
                    label="⬇️ Download Excel Report",
                    data=report_buffer,
                    file_name=f"reconciliation_report_{pd.Timestamp.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
                st.success("Report generated successfully!")
        
    except Exception as e:
        st.error(f"❌ Error processing files: {str(e)}")
        st.exception(e)

else:
    st.info("👆 Please upload both reconciliation Excel and invoice files to begin")
    
    # Display help
    with st.expander("ℹ️ How to use this tool"):
        st.markdown("""
        ### Step 1: Upload Files
        - **Reconciliation Excel**: Upload your reconciliation spreadsheet containing product types, denominations, quantities, and amounts
        - **Invoice File**: Upload the invoice in any supported format (XML, Excel, PDF, or Image)
        
        ### Step 2: Configure Settings
        Use the sidebar to adjust:
        - **VAT Tolerance**: Allow small differences in VAT calculations
        - **Amount Tolerance**: Allow small rounding differences in amounts
        - **Fuzzy Matching**: Enable flexible product name matching
        - **Compare After Discount**: Include discount columns in comparison
        
        ### Step 3: Review Results
        - Check the extracted data from both files
        - Review the comparison table with color-coded matches/mismatches
        - Verify totals comparison
        - Download the detailed Excel report
        
        ### Supported File Formats
        - **XML**: Electronic invoice XML files
        - **Excel**: .xlsx, .xls spreadsheets
        - **PDF**: Both text-based and scanned PDFs (with OCR)
        - **Images**: .png, .jpg, .jpeg (with OCR for Vietnamese text)
        """)
