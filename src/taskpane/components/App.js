import * as React from "react";
import PropTypes from "prop-types";
import Header from "./Header";
import Progress from "./Progress";
import Body from "./Body";

export default class App extends React.Component {
  constructor(props, context) {
    super(props, context);
  }

  render() {
    const { title, isOfficeInitialized } = this.props;

    if (!isOfficeInitialized) {
      return (
        <Progress
          title={title}
          logo={"./../../../assets/logo-filled.png"}
          message="Please sideload your addin to see app body."
        />
      );
    }

    return (
      <div className="ms-welcome">
        <Header title={title} />
        <Body />
      </div>
    );
  }
}

App.propTypes = {
  title: PropTypes.string,
  isOfficeInitialized: PropTypes.bool,
};
