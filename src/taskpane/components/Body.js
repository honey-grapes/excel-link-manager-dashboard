import React, { useEffect, useState } from "react";
import ItemList from "./ItemList";
import Formulas from "./Formulas";
import { DefaultButton, Dropdown, Pivot, PivotItem } from "@fluentui/react";

const Body = () => {
  const [selected, setSelected] = useState("");
  const [formulaSelected, setFormulaSelected] = useState("");
  const [precedents, setPrecedents] = useState([]);
  const [dependents, setDependents] = useState([]);
  const [history, setHistory] = useState([]);

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

  const onClickFetchList = async () => {
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
  };

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
            <Formulas formulas={formulaSelected} selected={selected} />
          </span>
        </div>
      </div>
      <div className="list-body">
        <Pivot styles={pivotStyles} onLinkClick={onClickFetchList}>
          <PivotItem key="precedents" headerText="Fetch Precedents" itemKey="precedents">
            <ItemList items={precedents} />
          </PivotItem>
          <PivotItem key="dependents" headerText="Fetch Dependents" itemKey="dependents">
            <ItemList items={precedents} />
          </PivotItem>
          <PivotItem key="offset" headerText="Fetch Offset" itemKey="offset">
            <div className="explanation">Currently only works for single cell range offset</div>
            <ItemList items={precedents} />
          </PivotItem>
          <PivotItem key="history" headerText="History" itemKey="history">
            <ItemList items={history} />
          </PivotItem>
        </Pivot>
      </div>
    </div>
  );
};

export default Body;
