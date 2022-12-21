import { BadRequestException, Injectable } from '@nestjs/common';
import { Workbook, Worksheet, Cell } from 'exceljs';
import * as tmp from 'tmp';

@Injectable()
export class ExcelService {
  async downloadExcel() {
    let data = [
      {
        name: 'user1',
        email: 'user1@gmail.com',
      },
      {
        name: 'user2',
        email: 'user2@gmail.com',
      },
      {
        name: 'user3',
        email: 'user3@gmail.com',
      },
    ];

    let rows = [];

    data.forEach((item) => {
      rows.push(Object.values(item));
      console.log(Object.values(item));
    });

    let workbook = new Workbook();

    let sheet = workbook.addWorksheet('sheet1');
    rows.unshift(Object.keys(data[0]));

    // sheet.addRows(rows);
    var rowValues = [];
    rowValues[1] = 4;
    rowValues[5] = 'Kyle';
    rowValues[9] = new Date();

    // insert new row and return as row object
    sheet.insertRow(1, rowValues);
    sheet.insertRow(2, rowValues);

    let File = await new Promise((resolve, reject) => {
      tmp.file(
        {
          discardDescriptor: true,
          prefix: 'MyExcel',
          postfix: '.xlsx',
          mode: parseInt('0600', 8),
        },
        async (err, file) => {
          if (err) {
            throw new BadRequestException(err);
          }
          workbook.xlsx
            .writeFile(file)
            .then((_) => {
              resolve(file);
            })
            .catch((err) => {
              throw new BadRequestException(err);
            });
        },
      );
    });

    return File;
  }

  async readFileExcel(file: Express.Multer.File) {
    // read from a file
    const workbook = new Workbook();
    await workbook.xlsx.load(file.buffer);
    // ... use workbook
    var worksheet = workbook.getWorksheet('口腔訓練');
    // worksheet.eachRow({ includeEmpty: true }, function (row, rowNumber) {
    //   console.log('Row ' + rowNumber + ' = ' + JSON.stringify(row.values));
    // });

    const firstRow = this.getCellByName(worksheet, '日付');

    const cell = this.getCellByName(worksheet, 'AM軽度氏名');
    cell.value = 'new value';

    firstRow.fill = {
      type: 'gradient',
      gradient: 'angle',
      degree: 0,
      stops: [
        { position: 0, color: { argb: 'FF0000FF' } },
        { position: 0.5, color: { argb: 'FFFFFFFF' } },
        { position: 1, color: { argb: 'FF0000FF' } },
      ],
    };

    let File = await new Promise((resolve, reject) => {
      tmp.file(
        {
          discardDescriptor: true,
          prefix: 'TestExcelPostMan',
          postfix: '.xlsx',
          mode: parseInt('0600', 8),
        },
        async (err, file) => {
          if (err) {
            throw new BadRequestException(err);
          }
          workbook.xlsx
            .writeFile(file)
            .then((_) => {
              resolve(file);
            })
            .catch((err) => {
              throw new BadRequestException(err);
            });
        },
      );
    });

    return File;
  }

  private getCellByName(worksheet: Worksheet, name: string): Cell {
    let match;
    worksheet.eachRow((row) =>
      row.eachCell((cell) => {
        if (cell.names.find((n) => n === name)) {
          match = cell;
        }
      }),
    );
    return match;
  }
}
