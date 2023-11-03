
<script>
  import { onMount } from 'svelte';
  import { writable } from 'svelte/store';
  // <img src="/slides_png/slide_22.png" alt="slide_22" width="1000" height="500">

  const width = writable(0);
  const height = writable(0);
  onMount(() => {
    width.set(window.innerWidth);
    height.set(window.innerHeight);
    document.addEventListener('keydown', function(event) {
      //Check if the pressed key is the left arrow key
      if (event.key === 'ArrowLeft' || event.key === 'PageUp') {
        // Navigate to the desired URL when the left arrow key is pressed
        window.location.href = '/slide_21'; // Replace with your desired URL
      }
      //Check if the pressed key is the right arrow key
      else if (event.key === 'ArrowRight' || event.key === 'PageDown') {
        // Navigate to the desired URL when the right arrow key is pressed
        window.location.href = '/slide_23'; // Replace with your desired URL
      }
    });
  });
</script>



<img src="/slides_png/slide_22.png" alt="slide_22" width="{$width}">
              
