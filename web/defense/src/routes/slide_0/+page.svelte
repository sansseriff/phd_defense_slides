
<script>
  import { onMount } from 'svelte';
  import { writable } from 'svelte/store';
  // <img src="/src/slides_png/slide_0.png" alt="slide_0" width="1000" height="500">

  const width = writable(0);
  const height = writable(0);
  onMount(() => {
    width.set(window.innerWidth);
    height.set(window.innerHeight);
    document.addEventListener('keydown', function(event) {
      //Check if the pressed key is the left arrow key
      if (event.key === 'ArrowLeft') {
        // Navigate to the desired URL when the left arrow key is pressed
        window.location.href = '/slide_0'; // Replace with your desired URL
      }
      //Check if the pressed key is the right arrow key
      else if (event.key === 'ArrowRight') {
        // Navigate to the desired URL when the right arrow key is pressed
        window.location.href = '/slide_1'; // Replace with your desired URL
      }
    });
  });
</script>


<img src="/src/slides_png/slide_0.png" alt="slide_0" width="{$width}" height="{$height}">
              
