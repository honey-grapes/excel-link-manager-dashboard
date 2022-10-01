import React from "react";
import { DetailsList, SelectionMode } from "@fluentui/react/lib/DetailsList";
import PropTypes from "prop-types";
import { toast } from "react-toastify";

const Formulas = ({ formulas }) => {
  //Custom styles and props for DetailsList group header
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
    const handleToggleCollapse = () => {
      const { onToggleCollapse, group } = headerProps;
      if (formulas.length > 0) {
        onToggleCollapse(group);
      } else {
        toast("Cell range is empty");
      }
    };

    return (
      <span>
        {defaultRender({
          ...headerProps,
          expandButtonIcon: "Add",
          onToggleCollapse: handleToggleCollapse,
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

  let groupCount = 0;
  try {
    groupCount = formulas.length;
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
      minWidth: 10,
      maxWidth: 20,
      isResizable: true,
    },
    {
      key: "column2",
      name: "Row",
      fieldName: "row",
      minWidth: 10,
      maxWidth: 20,
      isResizable: true,
    },
    {
      key: "column3",
      name: "Formula",
      fieldName: "formula",
      minWidth: 50,
      maxWidth: 80,
      isResizable: true,
    },
  ];

  return (
    <DetailsList
      items={formulas}
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
