
<script lang='ts'>
  import { onMount } from 'svelte';
  import { writable } from 'svelte/store';
  import { goto } from '$app/navigation';
  
  // <img src="/slides_png/slide_0.png" alt="slide_0" width="1000" height="500">

  // let connection: WebSocket | null = null;
  // let state: any = null;
  // const {connection, state} = 
  
  

  const width = writable(0);
  const height = writable(0);
  onMount(() => {
    
    width.set(window.innerWidth);
    height.set(window.innerHeight);
    document.addEventListener('keydown', function(event) {
      //Check if the pressed key is the left arrow key
      if (event.key === 'ArrowLeft' || event.key === 'PageUp') {
        // Navigate to the desired URL when the left arrow key is pressed
        // window.location.href = '/slide_0'; // Replace with your desired URL
        goto('/slide_0')
      }
      //Check if the pressed key is the right arrow key
      else if (event.key === 'ArrowRight' || event.key === 'PageDown') {
        // Navigate to the desired URL when the right arrow key is pressed
        // window.location.href = '/slide_1'; // Replace with your desired URL
        goto('/slide_1');
      }
    });
  });
</script>



<img src="/slides_png/slide_0.png" alt="slide_0" width="{$width}">
              
