import * as React from 'react';
import { Component } from 'react';

interface IState {
  count: number;
}

export default class MyWebPart extends Component<{}, IState> {
  constructor(props: {}) {
    super(props);
    this.state = {
      count: 0
    };
  }

  increment = () => {
    this.setState({ count: this.state.count + 1 });
  }

  decrement = () => {
    this.setState({ count: this.state.count - 1 });
  }

  render() {
    return (
      <div style={{ textAlign: 'center', marginTop: '50px' }}>
        <h2>Simple Counter</h2>
        <h3>Count: {this.state.count}</h3>
        <button onClick={this.increment}>Increment</button>
        <button onClick={this.decrement} style={{ marginLeft: '10px' }}>Decrement</button>
      </div>
    );
  }
}