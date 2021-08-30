/**
 * Class to setup the initialization process of the app. 
 */
import fs from 'fs';
import path from 'path';

import ExcelService from './services/parseExcel';

class InitApp {

    private EXCEL_FILE: string;

    /**
     * Constructor
     * @param {string}
     */
    constructor(file: string) {
        this.EXCEL_FILE = file;
        this.validateFile();
        this.startParsing();

    }

    /**
     * Method to validate the file provided for parsing.
     * @returns {void}
     */
    private validateFile(): void {
        if (!fs.existsSync(this.EXCEL_FILE)) {
            console.log(`File with path ${this.EXCEL_FILE} doesn't exist.`);
        }
    }

    /**
     * Method to the parsing of the excel workbook.
     * @returns {void}
     */
    private startParsing(): void {
        new ExcelService().startProcessingOfExcel(this.EXCEL_FILE);
    }

}

const init = new InitApp(
    path.join(
        __dirname,
        '..',
        'data',
        'Data Set.xlsx'
    )
);