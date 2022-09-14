import React, { useEffect, useState } from "react";
import ItemList from "./ItemList";
import { DefaultButton } from "@fluentui/react";

const Body = () => {
  const [selected, setSelected] = useState("");
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
        context.workbook.worksheets.onSelectionChanged.add(onSelectionChange);
      });
    };
    fetchSelection();
  }, []);

  //Update and display selection range each time the range changes
  const onSelectionChange = async () => {
    await Excel.run(async (context) => {
      //Get and set selection
      const range = context.workbook.getSelectedRange();
      range.load("address");
      await context.sync();
      setSelected(range.address);
    });
  };

  const onClick = async () => {
    console.log("bro");
  };

  return (
    <div>
      <h1>{selected}</h1>
      <DefaultButton className="ms-welcome__action" onClick={onClick}>
        Fetch Precedent
      </DefaultButton>
      <ItemList />
    </div>
  );
};

export default Body;
