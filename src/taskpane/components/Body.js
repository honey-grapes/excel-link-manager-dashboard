import React, { useEffect, useState } from "react";
import ItemList from "./ItemList";
import { DefaultButton, FocusZone } from "@fluentui/react";

const Body = () => {
  const [selected, setSelected] = useState("");
  const [precedents, setPrecedents] = useState([]);
  //Fire once when the component mounts to retrieve the initial selected range
  //and keep track of selection change
  useEffect(() => {
    const fetchSelection = async () => {
      await Excel.run(async (context) => {
        //Get and set initial selection
        const range = context.workbook.getSelectedRange();
        range.load("address");
        await context.sync();
        setSelected(range.address);

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
      await context.sync();
      setSelected(range.address);
    });
  };

  const onClickFetchList = async () => {
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
  };

  return (
    <div>
      <h1>{selected}</h1>
      <DefaultButton className="ms-welcome__action" onClick={onClickFetchList}>
        Fetch Precedent
      </DefaultButton>
      <FocusZone>
        <ItemList items={precedents} />
      </FocusZone>
    </div>
  );
};

export default Body;
