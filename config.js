const SCRAP_SHEET_URL = 'https://docs.google.com/spreadsheets/d/1C1MeLvv2PrOEQsqmXQplZt1xLuD5Z4KUGOrnGjYrh_A/gviz/tq?tqx=out:csv&sheet=Scrap';
const WORKFILTER_SHEET_URL = 'https://docs.google.com/spreadsheets/d/1C-_TRQcz2w0ia9bF9BfFcln5X_dX8T_cNfVUI-oU8ho/gviz/tq?tqx=out:csv&sheet=WorkFilter';
const EMPLOYEE_SHEET_URL = 'https://docs.google.com/spreadsheets/d/1eqVoLsZxGguEbRCC5rdI4iMVtQ7CK4T3uXRdx8zE3uw/gviz/tq?tqx=out:csv&sheet=EmployeeWeb';
const WARRANTY_SHEET_URL = 'https://docs.google.com/spreadsheets/d/19WWSESBcerEUlEDzM4lu3cglPoTHOWUl4P0prexqYpc/gviz/tq?tqx=out:csv&sheet=WarrantyData';
const TECHNICIAN_SHEET_URL = 'https://docs.google.com/spreadsheets/d/1EcP7N1Yr6fAzl17ItoqPstaxrIIf1Aly/gviz/tq?tqx=out:csv&sheet=Data';
const PARTS_SHEET_URL = 'https://docs.google.com/spreadsheets/d/1C-_TRQcz2w0ia9bF9BfFcln5X_dX8T_cNfVUI-oU8ho/gviz/tq?tqx=out:csv&sheet=Parts';

const GAS_API_URL = 'https://script.google.com/macros/s/AKfycbwq_ocJV50pVeUhuYJ5NJ114PPLDzCXuelrkYP5Mz33krtkbFp7VP5pzQEuzXnrGlDyXA/exec';

const COLUMNS = [
    { header: '<input type="checkbox" id="selectAll" onclick="toggleAllCheckboxes(this)">', source: 'C', key: 'checkbox' },
    { header: 'Status', source: 'C', key: 'status' },
    { header: 'Work Order', source: 'S', key: 'work order' },
    { header: 'Material', source: 'S', key: 'Spare Part Code' },
    { header: 'Spare Part Name', source: 'S', key: 'Spare Part Name' },
    { header: 'Codeซาก', source: 'S', key: 'old material code' },
    { header: 'Person', source: 'P', key: 'Person' },
    { header: 'Qty', source: 'S', key: 'qty' },
    { header: 'Serial Number', source: 'W', key: 'Serial Number' },
    { header: 'Store Code', source: 'W', key: 'Store Code' },
    { header: 'Store Name', source: 'W', key: 'Store Name' },
    { header: 'รหัสช่าง', source: 'S', key: 'รหัสช่าง' },
    { header: 'ชื่อช่าง', source: 'S', key: 'ชื่อช่าง' },
    { header: 'Mobile', source: 'T', key: 'Phone' },
    { header: 'Plant', source: 'S', key: 'plant' },
    { header: 'Product', source: 'W', key: 'Product' }
];

const PLANT_MAPPING = {
    '301': 'นวนคร', '303': 'SA-สงขลา', '304': 'พระราม 3', '305': 'ราชบุรี', '307': 'สุราษฎร์',
    '309': 'นครราชสีมา', '310': 'SA-อุดร 1', '311': 'ศรีราชา', '312': 'พิษณุโลก', '313': 'ภูเก็ต',
    '315': 'SA-อยุธยา', '319': 'ขอนแก่น', '318': 'คลังวัตถุดิบ', '320': 'ลำปาง', '322': 'SA-อุดร 2',
    '323': 'SA-ลำลูกกา', '324': 'SA-ปัตตานี', '326': 'วิภาวดี 62', '330': 'ประเวศ', '362': 'SA-หนองแขม',
    '363': 'SA-ปากเกร็ด', '364': 'SA-บางบัวทอง', '365': 'SA-สมุทรปราการ', '366': 'เชียงใหม่',
    '367': 'SA-ฉะเชิงเทรา', '368': 'SA-ร้อยเอ็ด', '369': 'ระยอง'
};

const BOOKING_COLUMNS = [
    { header: '<input type="checkbox" id="selectAllBooking" onclick="toggleAllBookingCheckboxes(this)">', key: 'checkbox' },
    { header: 'Status', key: 'CustomStatus' },
    { header: 'ใบจองรถ', key: 'Booking Slip' },
    { header: 'Work Order', key: 'Work Order' },
    { header: 'Material', key: 'Spare Part Code' },
    { header: 'Spare Part Name', key: 'Spare Part Name' },
    { header: 'Codeซาก', key: 'Old Material Code' },
    { header: 'Qty', key: 'Qty' },
    { header: 'Serial Number', key: 'Serial Number' },
    { header: 'Store Code', key: 'Store Code' },
    { header: 'Store Name', key: 'Store Name' },
    { header: 'Claim Receiver', key: 'Claim Receiver' },
    { header: 'Plantcenter', key: 'Plantcenter' },
    { header: 'Plant', key: 'Plant' },
    { header: 'Product', key: 'Product' },
    { header: 'วันที่จองรถ', key: 'Booking Date' },
];

const SUPPLIER_COLUMNS = [
    { header: '<input type="checkbox" id="selectAllSupplier" onclick="toggleAllSupplierCheckboxes(this)">', key: 'checkbox' },
    { header: 'Status', key: 'CustomStatus' },
    { header: 'ใบจองรถ', key: 'Booking Slip' },
    { header: 'Work Order', key: 'Work Order' },
    { header: 'Material', key: 'Spare Part Code' },
    { header: 'Spare Part Name', key: 'Spare Part Name' },
    { header: 'Codeซาก', key: 'Old Material Code' },
    { header: 'Qty', key: 'Qty' },
    { header: 'Serial Number', key: 'Serial Number' },
    { header: 'Store Code', key: 'Store Code' },
    { header: 'Store Name', key: 'Store Name' },
    { header: 'Claim Receiver', key: 'Claim Receiver' },
    { header: 'รหัสช่าง', key: 'รหัสช่าง' },
    { header: 'ชื่อช่าง', key: 'ชื่อช่าง' },
    { header: 'Mobile', key: 'Mobile' },
    { header: 'Plantcenter', key: 'Plantcenter' },
    { header: 'Plant', key: 'Plant' },
    { header: 'Product', key: 'Product' },
    { header: 'Warranty Action', key: 'Warranty Action' },
    { header: 'Recorder', key: 'Recorder' },
    { header: 'Recripte', key: 'Recripte' },
    { header: 'Recripte Date', key: 'RecripteDate' },
    { header: 'Timestamp', key: 'Timestamp' },
    { header: 'วันที่จองรถ', key: 'Booking Date' },
    { header: 'Claim Date', key: 'Claim Date' },
    { header: 'ClaimSup', key: 'ClaimSup' }
];

let allEmployees = [];
let fullData = [];
let displayedData = [];
let globalBookingData = [];
let currentPage = 1;
const ITEMS_PER_PAGE = 20;