/**
 * Service to parse the excel.
 */
import * as XLSX from 'xlsx';

export default class ExcelService {

    private workBook: XLSX.WorkBook;
    private sheet_names: string[];
    private sheetRecordsInJson: any[];
    private filteredRecordsInJson: any[];
    
    /**
     * method to validate the conditions and filter the results.
     * @returns {Promise<void>}
     */
    private async validatingConditionsAndFilteringRecords(): Promise<void> {
        this.filteredRecordsInJson = this.sheetRecordsInJson.filter((value: any) => {
            if (value['Record Type'] === 'Keyword' && parseFloat(value['ACoS']) > 50) {
                value['Max Bid'] = parseFloat(value['Max Bid']) * 0.9;
                return value;
            }
        });
    }

    /**
     * Method to start the processing of sheet. 
     * @param {string} file
     * @returns {Promise<void>}
     */
    public async startProcessingOfExcel(file: string): Promise<void> {
        console.log(`Starting the processing of excel work book ${file}.`);
        try {
            this.workBook = XLSX.readFile(file);
            this.sheet_names = this.workBook.SheetNames;
            if (this.sheet_names.indexOf('Sponsored Products Campaigns') >= 0) {
                this.sheetRecordsInJson = XLSX.utils.sheet_to_json(
                    this.workBook.Sheets[this.sheet_names[0]]
                );
                await this.validatingConditionsAndFilteringRecords();
                const filtered_sheet_name: string = `filtered_data-${new Date().getTime()}`;
                XLSX.utils.book_append_sheet(
                    this.workBook, 
                    XLSX.utils.json_to_sheet(this.filteredRecordsInJson), 
                    filtered_sheet_name
                );
                XLSX.writeFile(this.workBook, file);
                console.log(`Filtering of data done and added in new sheet named with ` +
                        `${filtered_sheet_name} in the existing workbook located at ${file}.`);
            } else {
                console.log(`Sheet doesn't exist with name 'Sponsored Products Campaigns'.` +
                        ` Exiting the process.`);
            }
        } catch (err) {
            console.log(`Error encounter while parsing ${err}.`);
        }
    }

}