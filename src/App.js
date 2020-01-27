import React from 'react';
import logo from './logo.svg';
import './App.css';
import { render } from '@testing-library/react';
/* import XLSX from 'xlsx'; */
import ExcelJS from 'exceljs'
import { saveAs } from 'file-saver';
function App() {

  var data = {
    title: 'report title',
    groups: [
      {
        title: 'group1',
        aggrigation: {
          total: '52'
        },
        utilisitaion: [
          {
            'a@merck.com': {
              aggrigation: {
                total: '21'
              },
              dates: {
                '2020/01/20': 1,
                '2020/01/21': 1,
                '2020/01/22': 1,
                '2020/01/23': 4,
                '2020/01/24': 1,
                '2020/01/25': 6,
                '2020/01/26': 7,
              }
            },
            'b@merck.com': {
              aggrigation: {
                total: '31'
              },
              dates: {
                '2020/01/20': 2,
                '2020/01/21': 3,
                '2020/01/22': 4,
                '2020/01/23': 4,
                '2020/01/24': 5,
                '2020/01/25': 6,
                '2020/01/26': 7,
              }
            }
          }
        ]
      },
      {
        title: 'group2',
        aggrigation: {
          total: '38'
        },
        utilisitaion: [
          {
            'c@merck.com': {
              aggrigation: {
                total: '7'
              },
              dates: {
                '2020/01/20': 1,
                '2020/01/21': 1,
                '2020/01/22': 1,
                '2020/01/23': 1,
                '2020/01/24': 1,
                '2020/01/25': 1,
                '2020/01/26': 1,
              }
            },
            'd@merck.com': {
              aggrigation: {
                total: '31'
              },
              dates: {
                '2020/01/20': 2,
                '2020/01/21': 3,
                '2020/01/22': 4,
                '2020/01/23': 4,
                '2020/01/24': 5,
                '2020/01/25': 6,
                '2020/01/26': 7,
              }
            }
          }
        ]
      }
    ]
  }

  var borderObj = {
    top: { style: 'thin', color: { argb: '000' } },
    left: { style: 'thin', color: { argb: '000' } },
    bottom: { style: 'thin', color: { argb: '000' } },
    right: { style: 'thin', color: { argb: '000' } }
  };
  async function createExcel() {

    var ws_data = [];
    let utilisitaion_keys = Object.keys(data.groups[0].utilisitaion[0]);
    let header = Object.keys(data.groups[0].utilisitaion[0][utilisitaion_keys[0]].dates)
    header.unshift('Email');
    header.push('total');
    //ws_data.push(header);
    for (let i = 0; i < data.groups.length; i++) {
      let group = [];
      group.push(data.groups[i].title);
      group.push(data.groups[i].aggrigation.total);
      ws_data.push(group);
      let groupProfiles = Object.keys(data.groups[i].utilisitaion[0]);
      groupProfiles.forEach((item, index) => {
        let groupProfDetails = [];
        groupProfDetails = Object.values(data.groups[i].utilisitaion[0][item].dates);
        groupProfDetails.unshift(item);
        groupProfDetails.push(data.groups[i].utilisitaion[0][item].aggrigation.total)
        ws_data.push(groupProfDetails);
      })
    }



    const wb = new ExcelJS.Workbook()

    const ws = wb.addWorksheet()

    let headerArray = []
    header.forEach((item) => {
      let obj = { header: item, key: item, width: 17 }
      headerArray.push(obj);
    })
    ws.columns = headerArray;


    ws_data.forEach(item => {
      ws.addRow(item)

    })

    ws.getRow(1).font = { size: 12, bold: true };



    ws.eachRow(function (row, rowNumber) {
      if (row._cells.length == 2) {
        row.eachCell(function (cell, colNumber) {
          row.getCell(colNumber).font = { color: { argb: "004e47cc" }, bold: true, size: 12 };

          row.getCell(colNumber).fill = {
            type: 'pattern',
            pattern: 'solid',
            fgColor: {
              argb: 'acff00'
            }
          };
          row.getCell(colNumber).border = borderObj;
        });
        row.splice(header.length, 1, row._cells[1]._value.value);
        row.getCell(header.length).fill = {
          type: 'pattern',
          pattern: 'solid',
          fgColor: {
            argb: '63ab7d'
          }

        };
        row.getCell(header.length).border = borderObj;
        row.commit();
        ws.mergeCells(rowNumber, 1, rowNumber, header.length - 1);

        // 
      }
      else {
        row.eachCell(function (cell, colNumber) {
          if (row.getCell(colNumber).value < 6) {
            row.getCell(colNumber).fill = {
              type: 'pattern',
              pattern: 'solid',
              fgColor: {
                argb: 'ff661a'
              }
            };
            row.getCell(colNumber).border = borderObj;
          }
        });
      }
    });



    const buf = await wb.xlsx.writeBuffer()

    saveAs(new Blob([buf]), 'abc.xlsx')
  }


  return (
    <div className="App">
      <div id="navbar"><span>Red Stapler - SheetJS </span></div>
      <div id="wrapper">

        <button id="button-a" onClick={createExcel}>Create Excel</button>
      </div>
    </div>
  );
}

export default App;
