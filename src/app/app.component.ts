import { Component } from '@angular/core';
import { FormBuilder, Validators } from '@angular/forms';
import { Workbook } from 'exceljs';
import * as fs from 'file-saver';
@Component({
  selector: 'my-app',
  templateUrl: './app.component.html',
  styleUrls: ['./app.component.css'],
})
export class AppComponent {
  name = 'Angular';
  list = [];
  typelist: any = [];
  NewMerge = [];
  firstselect = [];
  header = [];
  header1 = [];
  totalDueNewCardForm = this.fb.group({
    creditCardNumber: ['1234432112344321', Validators.required],
  });

  constructor(private fb: FormBuilder) {}
  ngOnInit() {
    this.firstselect = [
      { id: '', name: '' },
      { id: '', name: '' },
    ];
    this.list = [
      {
        eid: '',
        ename: '',
        esal: '',
        code: '',
      },
      {
        eid: '',
        ename: '',
        esal: '',
        code: '',
      },
      {
        eid: '',
        ename: '',
        esal: '',
        code: '',
      },
    ];
    this.header = [
      'Containers',
      'Pallets',
      'Racks',
      'Cylinders',
      'Bags',
      'Bins',
      'Crates',
      'Totes',
      'Kegs',
      'Drums',
      'Intermediate Bulk Containers (IBCs​)',
      'Equipment',
    ];

    // this.generateExcel(data, header);
    //alert('i ngOnInit!');
    this.NewMerge = [
      [
        'Flat rack container',
        'Open top container',
        'Open side container',
        'Refrigerated ISO',
        'ISO tanks',
        'Half height container',
        'Special purpose container',
        'Aoao',
        'sxcs',
      ],
      [
        'Wooden',
        'Plastic',
        'Metal',
        'Pallets',
        'Steel',
        'Aluminium',
        'Iron',
        'Stteel',
        'Allm',
        'Pipes',
        'Magnet',
        'sas',
      ],
      [
        'Selective Pallet Racks',
        'Drive-in Rack Systems',
        'Push-back Pallet Racks',
        'Pallet Flow Racks',
        'Carton Flow Racks',
        'Cantilever Racks',
        'AutoPilot',
        'ADDTAGS',
        'RackWell',
        'WillRacks',
        'steelracks',
      ],
      [
        'LPG domestic',
        'LPG export',
        'LPG auto gas tanks',
        'Industrial gases',
        'Non- refillables',
        'LPGHYDROGEN1',
        'LPGHYDROGEN2',
      ],
      [
        'Pillow bag',
        'Gusseted bag',
        'Standup Pounch',
        'Tetrahedral bag',
        'Quattro seal Bag',
        'Sachet',
        'Diaper bag (Delivery Bags)',
        'Ponytail bag',
        'Duffle top',
        'Spout top',
        'Open top',
        'Spout bottom',
        'Plain bottom',
        'Full bottom',
        'WhiteBags',
        'FIxBag',
        'Asks',
        'lack',
        'RedBag',
      ],
      ['Stacking Bin', 'Partition Bins', 'StackBins'],
      [
        'Foldable Plastic Crates',
        'Plastic Standard Crates',
        'Nestable or Stackable Plastic Crates',
        'Bottle Crates',
        'Open Wooden Crates',
        'Wooden Sheated Crates',
        'One-way and Two-way Crates',
      ],
      ['Steel Totes', 'Aluminum Totes', 'Stainless steel Totes', 'PVC'],
      [
        'Mini Keg',
        'Cornelius Keg',
        'Sixth Barrel Keg',
        'Quarter Barrel Keg',
        'Slim Quarter Keg',
        'Half Barrel Keg',
      ],
      [
        'Plastic Drums',
        'Steel Drums',
        'Fibre Drums',
        'AluminiumDrums',
        'WoodenDrums',
      ],
      [
        'Metal IBC',
        'Flexible IBC',
        'Rigid Plastic IBC',
        'Composite IBC',
        'Fibreboard IBC',
        'Wooden IBC',
        'MetaData',
      ],
      ['Machinary', 'Bench'],
      [
        'BeligiumChoclate',
        'Chocolate',
        'RedValvet',
        'Choco',
        'vennila',
        'Mango',
        'Pista',
        'Asset',
      ],
      ['television', 'TATA', 'YFC', 'YFB'],
      ['Chicken', 'Mutton', 'Fish'],
      ['Jeans', 'Tshirt', 'Pants', 'shirts', 'Whiteshirt'],
      ['Jet', 'Space', 'Speed'],
      ['Hoodiees', 'Sweeters', 'Hoody'],
      ['Cool'],
      ['BoxingBag'],
      [],
      [],
      [],
      [],
      [],
      [],
      [],
      [],
      [],
      [],
      [],
      [],
      ['lyf', 'smart'],
      ['no1'],
      ['Mental1'],
      [],
      [],
      [],
      ['NewRockets'],
      [],
      [],
      [],
      ['Pikachu'],
      ['plywood'],
      [],
      [],
    ];

    let obj: any;
    for (let k = 0; k <= this.NewMerge.length; k++) {
      let type = [];
      for (let i = 0; i <= this.NewMerge.length; i++) {
        if (this.NewMerge[i] != undefined) {
          if (this.NewMerge[i].length > 0) {
            if (this.NewMerge[i][k] != undefined) {
              obj = this.NewMerge[i];
              type.push(obj[k]);
            }else if (this.NewMerge[i][k] =='') {
              // obj = this.NewMerge[i];
              type.push('');
            }else{
              type.push('');
            }            
          }else{
            type.push('');
          }  
        }else{
          type.push('');
        }  
      }
      this.typelist.push(type);
    }

    console.log(this.typelist);
    // let type = [];
    // let newtypelist = [];
    // let a1 = [1, 2, 3, 4],
    //   a2 = ['a', 'b', 'c', 'd'],
    //   a3 = [9, 8, 7, 6],
    //   a4 = ['z', 'y', 'x', 'w'],
    //   result = [a1, a2, a3, a4].reduce((a, b) =>
    //     a.map((v, i) => v + b[i]));

    // console.log(result);
    

    // let array1: any = ['A', 'B', 'C'];

    // let array2: any = ['1', '2', '3', '4'];

    // console.log(array1.flatMap((d) =>console.log(d), array2.map((v) =>console.log(v))));
  }

  generateExcel(list, header) {
    let data: any = [];
    for (let i = 0; i < list.length; i++) {
      let arr = [list[i].eid, list[i].ename, list[i].esal, list[i].code];
      data.push(arr);
    }

    //Create workbook and worksheet
    let workbook = new Workbook();
    let worksheet1 = workbook.addWorksheet('Sheet3');
    let worksheet = workbook.addWorksheet('Sheet4');
    //Add Header Row
    let headerRow = worksheet.addRow(header);
    // let headerRow1 = worksheet1.addRow(this.header1);
    // Cell Style : Fill and Border
    headerRow.eachCell((cell, number) => {
      cell.fill = {
        type: 'pattern',
        pattern: 'solid',
        fgColor: { argb: 'FFFFFF00' },
        bgColor: { argb: 'FF0000FF' },
      };
      cell.border = {
        top: { style: 'thin' },
        left: { style: 'thin' },
        bottom: { style: 'thin' },
        right: { style: 'thin' },
      };
    });
    worksheet.getColumn(3).width = 50;
    var data1 = [
      [
        'Flat rack container',
        'Wooden',
        'Selective Pallet Racks',
        'LPG domestic',
        'Pillow bag',
        'Stacking Bin',
        'Foldable Plastic Crates',
        'Steel Totes',
        'Mini Keg',
        'Plastic Drums',
        'Metal IBC',
      ],
      [
        'Open top container',
        'Plastic',
        'Drive-in Rack Systems',
        'LPG export',
        'Gusseted bag',
        'Partition Bins',
        'Plastic Standard Crates',
        'Aluminum Totes',
        'Cornelius Keg',
        'Flexible IBC',
        '',
        '',
      ],
      [
        'Open side container',
        'Metal',
        'Push-back Pallet Racks',
        'LPG auto gas tanks',
        'Standup Pounch',
        '',
        'Nestable or Stackable Plastic Crates',
        'Stainless steel Totes',
        'Sixth Barrel Keg',
        'Fibre Drums',
        'Rigid Plastic IBC ',
        '',
        '',
      ],
      [
        'Refrigerated ISO',
        '',
        'Pallet Flow Racks',
        'Industrial gases ',
        'Tetrahedral bag',
        '',
        'Bottle Crates',
        'PVC',
        'Quarter Barrel Keg',
        '',
        'Composite IBC',
        '',
        '',
      ],
      [
        'ISO tanks',
        '',
        'Carton Flow Racks',
        'Non- refillables',
        'Quattro seal Ba',
        '',
        'Open Wooden Crates',
        '',
        'Slim Quarter Keg',
        '',
        'Fibreboard IBC ',
        '',
        '',
      ],
      [
        'Half height container',
        '',
        'Cantilever Racks',
        '',
        'Sachet',
        '',
        'Wooden Sheated Crates',
        '',
        'Half Barrel Keg',
        '',
        'Wooden IBC',
        '',
        '',
      ],
      [
        'Special purpose container',
        '',
        '',
        '',
        'Diaper bag (Delivery Bags)',
        '',
        'One-way and Two-way Crates',
        '',
        '',
        '',
        '',
        '',
        '',
      ],
      ['', '', '', '', 'Ponytail bag', '', '', '', '', '', '', '', ''],
      ['', '', '', '', 'Duffle top', '', '', '', '', '', '', '', ''],
      ['', '', '', '', 'Spout top', '', '', '', '', '', '', '', ''],
      ['', '', '', '', 'Open top', '', '', '', '', '', '', '', ''],
      ['', '', '', '', 'Spout bottom', '', '', '', '', '', '', '', ''],
      ['', '', '', '', 'Plain bottom', '', '', '', '', '', '', '', ''],
      ['', '', '', '', 'Full bottom', '', '', '', '', '', '', '', ''],
    ];
    // let row = worksheet.addRow(data1);
    data1.forEach((d) => {
      let row = worksheet.addRow(d);
    });
    data.forEach((e) => {
      let row1 = worksheet1.addRow(e);
    });
    // let row1 = worksheet.addRow(list);
    // list.forEach((element, index) => {
    for (let index = 2; index <= 500; index++) {
      worksheet1.getCell('A' + index).dataValidation = {
        type: 'list',
        allowBlank: true,
        formulae: ['=Sheet4!$A$1:$ZZ$1'],
      };
      worksheet1.getCell('B' + index).dataValidation = {
        type: 'list',
        allowBlank: true,
        formulae: [
          // '=OFFSET(Sheet4!$A$1, 1, MATCH(Sheet3!$A$1, Sheet4!$A$1:$O$1,0)-1, COUNTA(OFFSET(Sheet4!$A$1, 1, MATCH(Sheet3!$A$1, Sheet4!$A$1:$O$1,0)-1, 20)), 1)',
          // '=OFFSET(Sheet4!$A$1:$O$1, 1, MATCH(Sheet3!$A$1:$O$1, Sheet4!$A$1:$O$1,0)-1, COUNTA(OFFSET(Sheet4!$A$1:$O$1, 1, MATCH(Sheet3!$A$1:$O$1, Sheet4!$A$1:$O$1,0)-1, 20)), 1)'
          // '=OFFSET(Sheet4!$A$1, 1, MATCH($A$1, Sheet4!$A$1:$O$1,0)-1, COUNTA(OFFSET(Sheet4!$A$1, 1, MATCH($A$1, Sheet4!$A$1:$O$1,0)-1, 20)), 1)',
          '=OFFSET(Sheet4!$A$1, 1, MATCH($A$' +
            index +
            ', Sheet4!$A$1:$ZZ$1,0)-1, COUNTA(OFFSET(Sheet4!$A$1, 1, MATCH($A$' +
            index +
            ', Sheet4!$A$1:$ZZ$1,0)-1, 20)), 1)',
        ],
      };
    }
    // });
    //Generate Excel File with given name
    workbook.xlsx.writeBuffer().then((data) => {
      let blob = new Blob([data], {
        type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
      });
      fs.saveAs(blob, 'test.xlsx');
    });
  }
  onChangeCreditCardValue(event: string) {
    this.totalDueNewCardForm.get('creditCardNumber').patchValue(event);
  }
}
