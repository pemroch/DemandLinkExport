import { Component } from '@angular/core';
import * as XLSX from 'xlsx';

type AOA = any[][];

@Component({
    selector: 'app-root',
    templateUrl: './app.component.html',
    styleUrls: ['./app.component.css']
})
export class AppComponent {
    demandLink: any;
    dewarWorksheet: any;

    demandLinkData: any = {};
    exportName: string = '';

    wopts: XLSX.WritingOptions = { bookType: 'xlsx', type: 'array' };

    demandLinkChange() {
        const ref: any = document.getElementById('demand-link');
        const reader: FileReader = new FileReader();
        const target: DataTransfer = <DataTransfer>(ref);

        reader.onload = (e: any) => {
            const bstr: string = e.target.result;
            const wb: XLSX.WorkBook = XLSX.read(bstr, { type: 'binary' });
            const data = <AOA>(XLSX.utils.sheet_to_json(wb.Sheets[wb.SheetNames[0]], { header: 1 }));

            const StoreIDIndex = data[0].indexOf('StoreID');
            const FarmIDIndex = data[0].indexOf('FarmID');
            const ReplenSuggIndex = data[0].indexOf('ReplenSugg');

            this.demandLinkData = data.reduce((prev: any, curr: any) => {
                if (curr[ReplenSuggIndex]) {
                    if (!prev[curr[StoreIDIndex]]) {
                        prev[curr[StoreIDIndex]] = { [curr[FarmIDIndex]]: curr[ReplenSuggIndex] };
                    } else {
                        prev[curr[StoreIDIndex]][curr[FarmIDIndex]] = curr[ReplenSuggIndex];
                    }
                }
                return prev;
            }, {});

            console.log(this.demandLinkData)
        };

        reader.readAsBinaryString(target.files[0]);
    }

    dewarWorksheetChange() {
        const ref: any = document.getElementById('dewar-worksheet');
        const reader: FileReader = new FileReader();
        const target: DataTransfer = <DataTransfer>(ref);

        reader.onload = (e: any) => {
            const bstr: string = e.target.result;
            const readWb: XLSX.WorkBook = XLSX.read(bstr, { type: 'binary' });
            const rows = <AOA>(XLSX.utils.sheet_to_json(readWb.Sheets[readWb.SheetNames[0]], { header: 1 }));

            const newRows = rows.map((row, rowIndex) => {
                const newRow = [];

                for (let i = 0; i < row.length; i++) {
                    const storeNumber = row[2];

                    if (
                        (rowIndex < 5 && i > 12)
                        || (rowIndex > 4 && i < 13 && storeNumber)
                        || (rowIndex > 4 && i > 12 && storeNumber && !this.demandLinkData[storeNumber])
                    ) {
                        newRow.push(row[i] || null);
                    } else if (rowIndex > 4 && i > 12 && this.demandLinkData[storeNumber]) {;
                        newRow.push(this.demandLinkData[storeNumber][rows[3][i]] || null);
                    } else {
                        newRow.push(null);
                    }
                }

                return newRow;
            });

            const ws: XLSX.WorkSheet = XLSX.utils.aoa_to_sheet(newRows);
            const writeWb: XLSX.WorkBook = XLSX.utils.book_new();

            XLSX.utils.book_append_sheet(writeWb, ws, 'Sheet1');
            XLSX.writeFile(writeWb, this.exportName + '.xlsx');
        };

        reader.readAsBinaryString(target.files[0]);
    }
}
