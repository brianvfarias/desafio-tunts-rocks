import axios from 'axios';
import ExcelJS from 'exceljs';

// Fetching data from API
async function getCountries() {
  const countries = await axios
    .get('https://restcountries.com/v3.1/all')
    .then(res => res.data);
  return countries;
}

const countries = await getCountries();

// Filling worksheet with the data collected
const workbook = new ExcelJS.Workbook();

const countriesSheet = workbook.addWorksheet('Countries');

countriesSheet.columns = [
  {
    header: 'Name',
    key: 'name',
    width: 15
  },
  {
    header: 'Capital',
    key: 'capital',
    width: 15
  },
  {
    header: 'Area',
    key: 'area',
    width: 15
  },
  {
    header: 'Currencies',
    key: 'currencies',
    width: 15
  }
];

countries.forEach(({ name, capital, currencies, area }) => {
  countriesSheet.addRow({
    name: name.common,
    capital: capital ? JSON.stringify(capital[0]).replaceAll('"', '') : '-',
    area: area ? area : '-',
    currencies: currencies ? Object.keys(currencies).join(', ') : '-'
  });
});

// Editing and formatting the worksheet
countriesSheet.insertRow(0);

countriesSheet.mergeCells('A1:D1');
countriesSheet.getCell('A1').value = 'Countries List';
countriesSheet.getCell('A1').alignment = {
  vertical: 'middle',
  horizontal: 'center'
};
countriesSheet.getCell('A1').font = {
  size: 16,
  bold: true,
  color: { argb: '4F4F4F' }
};

const columnsTitleFormat = { size: 12, bold: true, color: { argb: '808080' } };

countriesSheet.getCell('A2').font = columnsTitleFormat;
countriesSheet.getCell('B2').font = columnsTitleFormat;
countriesSheet.getCell('C2').font = columnsTitleFormat;
countriesSheet.getCell('D2').font = columnsTitleFormat;

countriesSheet.getColumn(3).numFmt = '#.###,##';

// Creating worksheet file
countriesSheet.workbook.xlsx.writeFile('countries.xlsx');
