import * as React from "react";
import PropTypes from "prop-types";

export default class Header extends React.Component {
  render() {
    const { title } = this.props;
    return <section className="ms-header">{title}</section>;
  }
}

Header.propTypes = {
  title: PropTypes.string,
};
