import React from "react";
import { DetailsList, SelectionMode } from "@fluentui/react/lib/DetailsList";
import PropTypes from "prop-types";

const Formulas = ({ formulas, selected }) => {
  const handleRenderHeader = (headerProps, defaultRender) => {
    const headerContainerStyle = {
      height: "38px",
      ":hover": {
        backgroundColor: "white",
      },
      ":active": {
        backgroundColor: "white",
      },
    };
    const titleStyle = { color: "rgb(100,100,100)", fontSize: "14px", fontWeight: "400" };
    const headerCountStyle = { display: "none" };
    const expandStyle = {
      position: "absolute",
      height: "38px",
      top: 0,
      right: 0,
      left: "none",
      backgroundColor: "transparent",
      ":hover": {
        backgroundColor: "transparent",
      },
      ":active": {
        backgroundColor: "transparent",
      },
      i: {
        transform: "rotate(0deg)",
      },
    };
    return (
      <span>
        {defaultRender({
          ...headerProps,
          expandButtonIcon: "Add",
          styles: {
            title: titleStyle,
            headerCount: headerCountStyle,
            expand: expandStyle,
            groupHeaderContainer: headerContainerStyle,
          },
        })}
      </span>
    );
  };

  const handleRenderRow = (props, defaultRender) => {
    return (
      <span>
        {defaultRender({
          ...props,
          styles: {
            root: { width: "100%" },
          },
        })}
      </span>
    );
  };

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
  } catch (error) {
    formulaItem.push();
  }

  let groupCount = 0;
  try {
    groupCount = formulas.length * formulas[0].length;
  } catch (error) {
    groupCount = 0;
  }

  const groups = [
    {
      key: "group1",
      name: "Expand to see all formulas",
      startIndex: 0,
      count: groupCount,
      level: 0,
      isCollapsed: true,
    },
  ];
  const columns = [
    {
      key: "column1",
      name: "Col",
      fieldName: "col",
      minWidth: 5,
      maxWidth: 5,
      isResizable: true,
    },
    {
      key: "column2",
      name: "Row",
      fieldName: "row",
      minWidth: 5,
      maxWidth: 5,
      isResizable: true,
    },
    {
      key: "column3",
      name: "Formula",
      fieldName: "formula",
      minWidth: 80,
      maxWidth: 200,
      isResizable: true,
    },
  ];

  return (
    <DetailsList
      items={formulaItem}
      groups={groups}
      columns={columns}
      selectionMode={SelectionMode.none}
      isHeaderVisible={false}
      indentWidth="10"
      groupProps={{ onRenderHeader: handleRenderHeader, showEmptyGroups: true }}
      onRenderRow={handleRenderRow}
      styles={{
        root: {
          border: "1px solid rgb(219,219,219)",
          borderRadius: "2px",
        },
        contentWrapper: {
          minHeight: "38px",
          overflow: "hidden",
        },
      }}
    />
  );
};

export default Formulas;
Formulas.propTypes = {
  formulas: PropTypes.any,
  selected: PropTypes.string,
};
