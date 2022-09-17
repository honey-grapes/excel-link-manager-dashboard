import React from "react";
import { DetailsList, SelectionMode } from "@fluentui/react/lib/DetailsList";
import { ScrollablePane, ScrollbarVisibility } from "@fluentui/react";

const ItemList = ({ items }) => {
  const columns = [
    {
      key: "column1",
      name: "Name",
      fieldName: "name",
      minWidth: 20,
      maxWidth: 200,
      isResizable: true,
      isModalSelection: false,
      styleHeader: "dataListHeader",
    },
    {
      key: "column2",
      name: "Data Preview",
      fieldName: "data",
      minWidth: 20,
      maxWidth: 200,
      isResizable: true,
      isModalSelection: false,
      styleHeader: "dataListHeader",
    },
  ];

  return (
    <div style={{ position: "relative", height: "300px", width: "80%" }}>
      <ScrollablePane scrollbarVisibility={ScrollbarVisibility.always}>
        <DetailsList items={items} columns={columns} isHeaderVisible={true} selectionMode={SelectionMode.none} />
      </ScrollablePane>
    </div>
  );
};

export default ItemList;
