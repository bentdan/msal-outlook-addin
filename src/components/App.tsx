import * as React from "react";
import { connect, ConnectedProps } from "react-redux";
import { AppState } from "src/react-app-env";

interface AppProps extends PropsFromRedux {
}

class App extends React.Component<AppProps> {
  constructor(props: AppProps) {
    super(props);
  }

  render() {
    return (
      <div>
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

export default connector(App);