<script>
  let myText = "";

  if (typeof document !== "undefined") {
    myText = "This is running in the browser.";
    console.log("This is running in the browser.");
  } else {
    myText = "This is running on the server.";
    console.log("This is running on the server.");
  }
  // Add an event listener for keydown events on the document
  document.addEventListener("keydown", function (event) {
    // Check if the pressed key is the left arrow key
    if (event.key === "ArrowLeft") {
      // Navigate to the desired URL when the left arrow key is pressed
      window.location.href = "/about"; // Replace with your desired URL
    }
    // Check if the pressed key is the right arrow key
    else if (event.key === "ArrowRight") {
      // Navigate to the desired URL when the right arrow key is pressed
      window.location.href = "/sverdle"; // Replace with your desired URL
    }
  });
</script>

<h1>This is my custom page</h1>

<a href="/sverdle">Sverdle</a>
