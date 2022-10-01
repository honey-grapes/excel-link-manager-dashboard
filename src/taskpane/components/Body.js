import React, { useEffect, useState, useCallback } from "react";
import ItemList from "./ItemList";
import Formulas from "./Formulas";
import { Pivot, PivotItem } from "@fluentui/react";

const Body = () => {
  /*===================================
  useState
  ===================================*/
  const [selected, setSelected] = useState("");
  const [formulas, setFormulaSelected] = useState([]);
  const [formulaList, setFormulaList] = useState([]);
  const [precedents, setPrecedents] = useState([]);
  const [dependents, setDependents] = useState([]);
  const [offsets, setOffsets] = useState([]);
  const [history, setHistory] = useState([]);

  /*===================================
  useEffect
  ===================================*/
  //Fire once when the component mounts to retrieve the initial selected range
  //and keep track of selection change
  useEffect(() => {
    const fetchSelection = async () => {
      await Excel.run(async (context) => {
        //Get and set initial selection
        const range = context.workbook.getSelectedRange();
        range.load("address");
        range.load("formulas");
        await context.sync();
        setSelected(range.address);
        setFormulaSelected(range.formulas);
        //Register and start tracking selection change
        context.workbook.onSelectionChanged.add(onSelectionChange);
      });
    };
    fetchSelection();
  }, []);

  //Fire everytime the formulas change for the selection to fetch
  //and process the formula so it is good for display
  useEffect(() => {
    const formulaItem = [];
    try {
      //Single cell selection (without ':')
      if (selected.indexOf(":") === -1) {
        if (formulas[0][0]) {
          const sheetAddressSplit = selected.replace(/!([^'])/g, "**$1").split("**");
          const cell = sheetAddressSplit[1];
          //Process cell
          let cellNumberStart = 1;
          for (let i = 1; i < cell.length; i++) {
            if (/[a-zA-Z]/.test(cell[i])) {
              continue;
            } else {
              cellNumberStart = i;
              break;
            }
          }
          const col = cell.slice(0, cellNumberStart);
          const row = parseInt(cell.slice(cellNumberStart));
          formulaItem.push({
            key: 1,
            col: col,
            row: row,
            formula: formulas[0][0],
          });
        }
      }
      //Multiple cell selection (with ':')
      else {
        //Read the start and end cell of the current selection so we can match
        //each formula to the corresponding cell
        const sheetAddressSplit = selected.replace(/!([^'])/g, "**$1").split("**");
        const startEndSplit = sheetAddressSplit[1].split(":");
        const startCell = startEndSplit[0];

        //Process start cell
        let cellNumberStart = 1;
        for (let i = 1; i < startCell.length; i++) {
          if (/[a-zA-Z]/.test(startCell[i])) {
            continue;
          } else {
            cellNumberStart = i;
            break;
          }
        }
        const startCellAlphabet = startCell.slice(0, cellNumberStart);
        const startRowInt = parseInt(startCell.slice(cellNumberStart));

        //Find looping through the number of rows and columns of the formula matrix
        //Rows = numbers and Columns = alphabets
        //
        //In order to loop through the columns and convert them into the
        //correct column letters such as A, AZ, XEF...etc:
        //  [1] Start looping from the number that corresponds to the Starting Cell's col letters
        //  [2] Convert incremented number back to column letter for display
        //The conversion can be done through mutiplication and division by the power of 27
        let startColInt = 0;
        let power = startCellAlphabet.length - 1;
        for (let i = 0; i < startCellAlphabet.length; i++) {
          startColInt += Math.pow(27, power - i) * (startCellAlphabet.charCodeAt(i) - 65 + 1);
        }
        //Start looping through the formulas matrix
        for (let c = 0; c < formulas[0].length; c++) {
          for (let r = 0; r < formulas.length; r++) {
            if (formulas[r][c]) {
              //Convert column number back to column letters
              let powerTwo = 2;
              let letter = "";
              let num = c + startColInt;
              while (power >= 0 && num > 0) {
                let div = Math.floor(num / Math.pow(27, powerTwo));
                if (div !== 0) {
                  letter += String.fromCharCode(65 + div - 1);
                  num -= div * Math.pow(27, powerTwo);
                }
                powerTwo -= 1;
              }

              formulaItem.push({
                key: r.toString() + c.toString(),
                col: letter,
                row: r + startRowInt,
                formula: formulas[r][c],
              });
            }
          }
        }
      }

      //Column header only added if there are formulaItems
      //If not, notify the user there is no formula to see (cell range is empty)
      if (formulaItem.length > 0) {
        formulaItem.unshift({
          key: 0,
          col: "C",
          row: "R",
          formula: "Formula",
        });
      }
    } catch (error) {
      formulaItem.push();
    }
    setFormulaList(formulaItem);
  }, [formulas]);

  /*===================================
  Event Handlers
  ===================================*/
  //Update and display selection range each time the range changes
  const onSelectionChange = async () => {
    await Excel.run(async (context) => {
      //Get and set current selection
      const range = context.workbook.getSelectedRange();
      range.load("address");
      range.load("formulas");
      await context.sync();
      setSelected(range.address);
      setFormulaSelected(range.formulas);
      setPrecedents([]);
    });
  };

  //Fetch Precedents
  const onClickFetchPre = useCallback(async () => {
    try {
      await Excel.run(async (context) => {
        //Get precedent cell ranges
        const range = context.workbook.getSelectedRange();
        const directPrecedents = range.getDirectPrecedents();
        directPrecedents.load("areas/items/address");
        await context.sync();

        //Format the addresses of precedents (string) into an array via splitting by commas
        //*** Note ***
        //There is an edge case is if ',' is the name of the worksheet
        //Since the name of the worksheet is always wrapped by '' if at least 1 character
        //is not alphanumerical, we can use regex to only split the address string by commas
        //without a '' wrapped around it. We can replace it with a character that Excel
        //prohibits in worksheet names such as \/?*[]. In this case, ** is used.
        let addressArray = [];
        for (let i = 0; i < directPrecedents.areas.items.length; i++) {
          const precString = directPrecedents.areas.items[i].address;
          const precArray = precString.replace(/([^']),/g, "$1**").split("**");
          addressArray.push(...precArray);
        }

        //Use a split loop to extract a preview of the data at the addresses of the precedents
        //to give the users a general sense of what data is stored in the precedent cells currently
        //*** Note ***
        //A split loop is used here because we should not use async context.sync() in every iteration
        //as it is resource intensive
        //Details: https://docs.microsoft.com/en-us/office/dev/add-ins/concepts/correlated-objects-pattern
        const previewDataArr = [];
        for (let i = 0; i < addressArray.length; i++) {
          //Use regex to split sheet name and address for data loading purpose
          const sheetAddressSplit = addressArray[i].replace(/!([^'])/g, "**$1").split("**");
          const sheetName = sheetAddressSplit[0].replace(/'/g, "");
          const sheet = context.workbook.worksheets.getItem(sheetName);
          const range = sheet.getRange(sheetAddressSplit[1]);
          range.load("values");
          previewDataArr.push(range);
        }
        //Load data in the entire preview array
        await context.sync();
        //Store preview data in a string less than 30 characters long and abbreviate the rest with "..."
        const previewStringArr = [];
        let totalLen = 0;
        for (let j = 0; j < previewDataArr.length; j++) {
          let tmp = [];
          for (let i = 0; i < previewDataArr[j].values.length; i++) {
            const curLen = previewDataArr[j].values[i][0].toString().length;
            //Plus 2 for the comma and space between each value
            if ((i != 0 && 30 - totalLen >= curLen + 2) || (i == 0 && 30 - totalLen >= curLen)) {
              totalLen += i == 0 ? curLen : curLen + 2;
              tmp.push(previewDataArr[j].values[i][0]);
            } else {
              tmp.push("...");
              break;
            }
          }
          previewStringArr.push(tmp.join(", "));
        }

        const resultArray = [];
        for (let i = 0; i < addressArray.length; i++) {
          resultArray.push({
            key: i,
            name: addressArray[i],
            data: previewStringArr[i],
          });
        }
        setPrecedents(resultArray);
      });
    } catch (error) {
      setPrecedents([]);
    }
  });

  //Fetch and process offset
  const onClickFetchOffset = useCallback(() => {
    const arr = [];
    arr.push({ key: 0, name: "dog", data: "bro" });
    setOffsets(arr);
  });

  //Handle different tabs in the pivot bar
  const handleLink = ({ props }) => {
    const { itemKey } = props;
    switch (itemKey) {
      case "precedents":
        onClickFetchPre();
        break;
      case "offset":
        onClickFetchOffset();
        break;
      case "dependents":
        console.log("dep");
        break;
      default:
        break;
    }

    //After each tab is clicked, record in history
    //History only records the 10 most recent actions
    if (itemKey != "history") {
      const count = history.length + 1;
      const curArr = history;
      if (count > 10) {
        curArr.pop();
      }

      const histId = Math.random().toString(36).substring(2);
      const act = itemKey.toString();
      curArr.unshift({
        key: histId,
        action: "Fetch " + act.charAt(0).toUpperCase() + act.slice(1),
        name: selected,
        data: selected,
      });
      setHistory(curArr);
    }
  };

  /*===================================
  Custom Style
  ===================================*/
  const pivotStyles = {
    root: {
      backgroundColor: "transparent",
    },
    link: {
      margin: "0",
      ":hover": {
        backgroundColor: "transparent",
        color: "#4e9668",
      },
      ":active": {
        backgroundColor: "transparent",
        color: "#4e9668",
      },
    },
    linkIsSelected: {
      margin: "0",
      color: "#4e9668",
      backgroundColor: "transparent",
      ":hover": {
        backgroundColor: "transparent",
        color: "#4e9668",
      },
      "::before": {
        backgroundColor: "#4e9668",
      },
      ":active": {
        backgroundColor: "transparent",
        color: "#4e9668",
      },
    },
  };

  return (
    <div>
      <div className="current-selection-header">Current Status</div>
      <div id="cur-select-body" className="current-selection-body">
        <div>
          <span className="current-selection-title">Selection </span>
          <span className="current-selection-value">{selected}</span>
        </div>
        <div className="current-selection-item">
          <span className="current-selection-title-formula">Formula </span>
          <span className="current-selection-value">
            <Formulas formulas={formulaList} selected={selected} />
          </span>
        </div>
      </div>
      <div className="list-body">
        <Pivot styles={pivotStyles} onLinkClick={handleLink}>
          <PivotItem key="precedents" headerText="Fetch Precedents" itemKey="precedents">
            <div className="direction">
              Go to <b>precedent</b> by clicking the corresponding line items
            </div>
            <ItemList items={precedents} listType="normal" />
          </PivotItem>
          <PivotItem key="dependents" headerText="Fetch Dependents" itemKey="dependents">
            <div className="direction">
              Go to <b>dependent</b> by clicking the corresponding line items
            </div>
            <ItemList items={dependents} listType="normal" />
          </PivotItem>
          <PivotItem key="offset" headerText="Fetch Offset" itemKey="offset">
            <div className="explanation">
              Currently only works for <b>single</b> cell range offset
            </div>
            <ItemList items={offsets} listType="offset" />
          </PivotItem>
          <PivotItem key="history" headerText="History" itemKey="history">
            <div className="direction">
              Go to each of the <b>10 most recent actions</b> by clicking the line item
            </div>
            <ItemList items={history} listType="history" />
          </PivotItem>
        </Pivot>
      </div>
    </div>
  );
};

export default Body;
