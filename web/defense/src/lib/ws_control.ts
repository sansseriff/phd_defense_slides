
// Function to establish a connection to the WebSocket server for a specific image
  export function connectToBackend() {

    // if (typeof window === 'undefined') {
    //   console.error('This function can only be run in a browser environment');
    //   return;
    // }
    let connection: WebSocket | null = null;

    let slide_state = {slide: 1}

    connection = new WebSocket(`ws://localhost:8000/ws/`);

    connection.onopen = (event) => {
      // Once the connection is open, start sending the chunks of the current image
      if (connection) {
        connection.send(JSON.stringify(slide_state));
      }
    };

    // Set the event handler for receiving messages
    connection.onmessage = (event) => console.log("message received: ", event.data);

    // Set the event handler for errors
    connection.onerror = (event) => {
      console.error('WebSocket error observed:', event);
    };


    // Set the event handler for closing the connection
    connection.onclose = (event) => {
    };

    return {connection, slide_state};
  }

