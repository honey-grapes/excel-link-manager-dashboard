import React from "react";
import { DetailsList, SelectionMode } from "@fluentui/react/lib/DetailsList";
import { ScrollablePane } from "@fluentui/react";

const ItemList = ({ items, listType }) => {
  let columns = [];
  switch (listType) {
    case "offset":
      columns = [
        {
          key: "column1",
          name: "Go to Range",
          fieldName: "name",
          minWidth: 20,
          maxWidth: 200,
          isResizable: true,
        },
        {
          key: "column2",
          name: "Snapshot",
          fieldName: "data",
          minWidth: 20,
          maxWidth: 200,
          isResizable: true,
        },
      ];
      break;
    case "history":
      columns = [
        {
          key: "column0",
          name: "Action",
          fieldName: "action",
          minWidth: 20,
          maxWidth: 100,
          isResizable: true,
        },
        {
          key: "column1",
          name: "Performed on",
          fieldName: "name",
          minWidth: 30,
          maxWidth: 150,
          isResizable: true,
        },
        {
          key: "column2",
          name: "Snapshot",
          fieldName: "data",
          minWidth: 30,
          maxWidth: 100,
          isResizable: true,
        },
      ];
      break;
    default:
      columns = [
        {
          key: "column1",
          name: "Go to Range",
          fieldName: "name",
          minWidth: 20,
          maxWidth: 200,
          isResizable: true,
        },
        {
          key: "column2",
          name: "Snapshot",
          fieldName: "data",
          minWidth: 20,
          maxWidth: 200,
          isResizable: true,
        },
      ];
  }

  const goToRange = (item) => {
    //Use regex to split sheet name and address for data loading purpose
    Excel.run(async (context) => {
      const sheetAddressSplit = item.name.replace(/!([^'])/g, "**$1").split("**");
      const sheetName = sheetAddressSplit[0].replace(/'/g, "");
      const sheet = context.workbook.worksheets.getItem(sheetName);
      const range = sheet.getRange(sheetAddressSplit[1]);

      range.select();
      await context.sync();
    });
  };

  return (
    <div style={{ position: "relative", height: "400px", width: "100%" }}>
      <ScrollablePane>
        <DetailsList
          items={items}
          columns={columns}
          isHeaderVisible={true}
          selectionMode={SelectionMode.none}
          onActiveItemChanged={goToRange}
        />
      </ScrollablePane>
    </div>
  );
};

export default ItemList;
