import { Component, ElementRef, Input, HostListener } from '@angular/core';
import * as XLSX from 'xlsx-style/dist/xlsx.full.min.js';
// declare var XLSX : any;
// import 'script-loader!../../node_modules/xlsx/dist/xlsx.full.min.js';
import * as FileSaver from 'file-saver';
import * as _ from 'lodash';


@Component({
    selector: 'app-export-to-excel',
    template : `
        <button class="btn btn-warning btn-sm">
            <i class="fa fa-file-excel-o" aria-hidden="true">&nbsp;</i>Export
        </button>
    `
})
export class ExportToExcelComponent {

    @Input() data: any ;

    sheetName = 'Test Sheet';
    workbook: XLSX.IWorkBook = {
        Sheets: {},
        SheetNames: [],
        Props: {}
    };
    ws: any;
    wbout: any;

    constructor(el: ElementRef) {}

    transformData (data: any): Array<any> {
        const dataNew: Array<any> = [];
        const keys_arr = [];
        _.forEach(data, function(json) {
            const arr = _.filter(json, function(value, index) {
                if (typeof value !== 'object') {
                    keys_arr.push( _.startCase(index));
                    return value;
                }
            });
            dataNew.push(arr);
        });
        dataNew.unshift(_.uniq(keys_arr));
        return dataNew;
    }

    @HostListener('click', ['$event']) onclick() {
        this.downloadExcel();
    }

    sheet_from_array_of_arrays(data) {
        console.log('data',data);
        const ws = {};
        const endCell = { c: 10000000, r: 10000000 };
        const startCell = { c: 0, r: 0 };
        const range = { s: endCell, e: startCell };

        const wscols = [];

        for (let R = 0; R < data.length; ++R) {
            for (let C = 0; C < data[R].length; ++C) {
                wscols.push({ wch: 20 });
                if (range.s.r > R) { range.s.r = R; }
                if (range.s.c > C) { range.s.c = C; }
                if (range.e.r < R) { range.e.r = R; }
                if (range.e.c < C) { range.e.c = C; }
                const cell = { v: data[R][C], t: 's' , s : {}};

                if (R === 0) {
                    cell.s = {
                        'font': {
                            'bold': true, 'sz': 13, 'alignment': {
                                'horizontal': 'center',
                                'vertical': 'center'
                            }
                        }
                    };
                }

                if (cell.v === null) {
                    continue;
                }

                const cell_ref = XLSX.utils.encode_cell({ c: C, r: R });
                if (typeof cell.v === 'number') {
                    cell.t = 'n';
                } else if (typeof cell.v === 'boolean') {
                    cell.t = 'b';
                } else {
                    cell.t = 's';
                }
                ws[cell_ref] = cell;
            }
        }
        ws['!cols'] = wscols;
        console.log('Worksheet goes here', ws);

        if (range.s.c < 10000000) {
            ws['!ref'] = XLSX.utils.encode_range(endCell, startCell);
        }
        return ws;
    }

    datenum(v, date1904?): any {
        if (date1904) {
            v += 1462;
        }
        const epoch = Date.parse(v);
        const dt: any = new Date(Date.UTC(1899, 11, 30));
        return (epoch - dt) / (24 * 60 * 60 * 1000);
    }

    generateExcelFile(): any {
        this.ws = this.sheet_from_array_of_arrays(this.transformData(this.data));
        this.workbook.SheetNames.push(this.sheetName);
        this.workbook.Sheets[this.sheetName] = this.ws;
        this.wbout = XLSX.write(this.workbook, { bookType: 'xlsx', type: 'binary' , cellStyles: true });
        return this.wbout;
    }

    s2ab(s: any): ArrayBuffer {
        const buf = new ArrayBuffer(s.length);
        const view = new Uint8Array(buf);
        for (let i = 0; i < s.length; ++i) {
            view[i] = s.charCodeAt(i) & 0xFF;
        }
        return buf;
    }

    downloadExcel(): void {
        FileSaver.saveAs(
            new Blob(
                [this.s2ab(this.generateExcelFile())],
                { type: 'application/octet-stream' }
            ),
            'test.xlsx'
        );
    }

}