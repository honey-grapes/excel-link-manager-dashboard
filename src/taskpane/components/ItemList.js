import * as React from "react";
import { FocusZone, FocusZoneDirection } from "@fluentui/react/lib/FocusZone";
import { List } from "@fluentui/react/lib/List";

const ItemList = () => {
  const items = [
    { key: 1, name: "file" },
    { key: 2, name: "file" },
  ];

  return (
    <FocusZone direction={FocusZoneDirection.vertical}>
      <div data-is-scrollable>
        <List items={items} />
      </div>
    </FocusZone>
  );
};

export default ItemList;
