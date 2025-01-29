import * as React from "react";
import { connect, ConnectedProps } from "react-redux";

interface MainProps extends PropsFromRedux {
  title: string;
}

class Main extends React.Component<MainProps> {
  constructor(props: MainProps) {
    super(props);
  }

  render() {
    return (
      <div>
        <h1>{this.props.title}</h1>
        <p>We are here!</p>
      </div>
    );
  }
}

// Map state and dispatch to props using redux
const mapState = (state: AppState) => ({
  isOfficeInitialized: state.officeReducers?.isOfficeInitialized,
});

const mapDispatch = {
};

const connector = connect(mapState, mapDispatch);
type PropsFromRedux = ConnectedProps<typeof connector>;

export default connector(Main);