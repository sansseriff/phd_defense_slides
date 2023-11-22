<script lang="ts">
  // import svelteLogo from "./assets/svelte.svg";
  // import viteLogo from "/vite.svg";
  // import Counter from "./lib/Counter.svelte";
  import { onMount } from "svelte";
  import { randomNormal } from "d3-random";
  import { diff, re, round } from "mathjs";
  import { tweened } from "svelte/motion";
  import { cubicOut } from "svelte/easing";
  import { tick } from "svelte";
  import { browser } from "$app/environment";

  const slider_value_1 = tweened(0, {
    duration: 2000,
    easing: cubicOut,
  });

  const slider_value_2 = tweened(0, {
    duration: 2000,
    easing: cubicOut,
  });

  // import * as bbokeh from "@bokeh/bokehjs"

  // console.log(bbokeh)

  interface JsonData {
    x: number[];
    y: number[];
  }

  let thing = false;

  interface WaveformData {
    sliding_pulse_1: Array<number>;
    sliding_pulse_2: Array<number>;
    existing_waveform: Array<number>;
    current_waveform: Array<number>;
  }

  interface Thing {
    waveform: WaveformData;
    sliding_1: Array<number>;
    sliding_2: Array<number>;
  }

  let sliding_pulse_1: Array<number>;
  let sliding_pulse_2: Array<number>;

  let jsonData;
  var scope = null;
  var scope_inset = null;
  // let slider_1: HTMLInputElement;
  // let slider_2: HTMLInputElement;
  // let output_1: HTMLElement;
  // let output_2: HTMLElement;
  // let slider_value_1: number
  // let slider_value_2: number

  let waveformData: WaveformData;
  let data: any;

  async function loadJsonData(path: string): Promise<JsonData> {
    const response = await fetch(path);
    return await response.json();
  }

  function rollArray<T>(arr: T[], n: number): T[] {
    arr.unshift(...arr.splice(-n));
    return arr;
  }

  function createRandomArray(n: number, p: number): number[] {
    const arr = new Array(n);
    const rand = randomNormal(0, p);
    for (let i = 0; i < n; i++) {
      arr[i] = rand();
    }
    return arr;
  }

  $: {
    if (thing) {
      let idx_1 = Math.round($slider_value_1);
      let idx_2 = Math.round($slider_value_2);

      let thing = update_graph(waveformData, idx_1, idx_2);

      data.data.y = thing.waveform.current_waveform;
      data.data.y_sliding_1 = thing.sliding_1;
      data.data.y_sliding_2 = thing.sliding_2;
      data.change.emit();
    } else {
      thing = true;
    }
  }

  async function setup() {
    const jsonData = await loadJsonData("/peacoq_pulse_extended.json");
    make_graph(jsonData);
  }

  // onMount(async () => {
  //   const jsonData = await loadJsonData();
  //   make_graph(jsonData);
  // });

  function update_graph(
    waveformData: WaveformData,
    idx_1: number,
    idx_2: number
  ): Thing {
    let sliding_pulse_1 = waveformData.sliding_pulse_1.slice();
    let sliding_pulse_2 = waveformData.sliding_pulse_2.slice();
    for (let i = 0; i < waveformData.existing_waveform.length; i++) {
      if (
        i - idx_1 - idx_2 < 0 ||
        i - idx_1 - idx_2 > waveformData.sliding_pulse_1.length
      ) {
        waveformData.current_waveform[i] = waveformData.existing_waveform[i];
      } else {
        waveformData.current_waveform[i] =
          waveformData.existing_waveform[i] +
          waveformData.sliding_pulse_1[i - idx_1] +
          waveformData.sliding_pulse_2[i - idx_2];

        sliding_pulse_1[i] = waveformData.sliding_pulse_1[i - idx_1];
        sliding_pulse_2[i] =
          waveformData.sliding_pulse_2[i - idx_2] +
          waveformData.sliding_pulse_1[i - idx_1];
      }
    }
    // waveformData.sliding_pulse_1 = sliding_pulse_1
    // waveformData.sliding_pulse_2 = sliding_pulse_2

    let thing = {
      waveform: waveformData,
      sliding_1: sliding_pulse_1,
      sliding_2: sliding_pulse_2,
    };

    return thing;
  }

  // function handle_sliders() {

  //   output_1.innerHTML = slider_1.value
  //   output_2.innerHTML = slider_2.value;

  //   var idx_1 = Math.round(parseFloat(slider_1.value));
  //   var idx_2 = Math.round(parseFloat(slider_2.value));
  //   let sliding_pulse_1: Array<number>;
  //   let sliding_pulse_2: Array<number>;
  //   [waveformData, sliding_pulse_1, sliding_pulse_2] = update_graph(waveformData, idx_1, idx_2);

  //   data.data.y = waveformData.current_waveform;
  //   data.data.y_sliding_1 = sliding_pulse_1;
  //   data.data.y_sliding_2 = sliding_pulse_2;
  //   data.change.emit();
  // }

  async function make_graph(json_data: JsonData) {
    waveformData = {
      sliding_pulse_1: json_data.y.slice(),
      sliding_pulse_2: json_data.y.slice(),
      existing_waveform: json_data.y.slice(),
      current_waveform: json_data.y.slice(),
    };

    let pulse_offset = createRandomArray(1300, 0.2);

    waveformData.existing_waveform.splice(
      -pulse_offset.length,
      pulse_offset.length
    );
    waveformData.existing_waveform.unshift(...pulse_offset);

    data = new Bokeh.ColumnDataSource({
      data: {
        x: json_data.x.map((value) => value - 3200 + 84),
        y: waveformData.current_waveform,
        y_stationary: waveformData.existing_waveform,
        y_sliding_1: waveformData.sliding_pulse_1,
        y_sliding_2: waveformData.sliding_pulse_2,
      },
    });

    const ToolTips = { index: "$index", "(x,y)": "($x,$y)" };
    var custom_tooltips = [
      ["X", "@x"],
      ["Y", "@y"],
    ];
    // var custom_hover = new Bokeh.HoverTool({
    //   tooltips: custom_tooltips,
    //   mode: "mouse",
    // });
    scope = new Bokeh.Plotting.figure({
      // title: "Example of random data",
      // tools: "pan,wheel_zoom,box_zoom,reset,save",
      tools: "",
      // toolbar: {logo: null},
      //tools: "",
      // sizing_mode: "stretch_both",
      sizing_mode: "stretch_width",
      height: 600,
      width: 2000,
      x_range: [-1000, 1000],
      y_range: [-80, 300],
    });

    scope.toolbar.logo = null;

    scope_inset = new Bokeh.Plotting.figure({
      tools: "",
      // sizing_mode: "stretch_width",
      height: 300,
      width: 300,
      x_range: [-1, 1],
      y_range: [150 - 100, 150 + 100],
    });
    scope_inset.toolbar.logo = null;

    // scope.add_tools(custom_hover);

    const band_1 = new Bokeh.Band({
      base: { field: "x" },
      lower: 0,
      upper: { field: "y_sliding_2" },
      source: data,
      fill_alpha: 0.3,
      line_alpha: 0,
      fill_color: "#f28c63",
    });

    const band_2 = new Bokeh.Band({
      base: { field: "x" },
      lower: { field: "y_sliding_2" },
      upper: { field: "y_sliding_1" },
      source: data,
      fill_alpha: 0.2,
      line_alpha: 0,
      fill_color: "#f542b0",
    });

    const line_1 = scope.line(
      { field: "x" },
      { field: "y" },
      {
        source: data,
        line_width: 4,
        line_color: "#000000",
      }
    );

    scope.add_layout(band_1);
    scope.add_layout(band_2);

    const line_inset = scope_inset.line(
      { field: "x" },
      { field: "y" },
      {
        source: data,
        line_width: 3,
        line_color: "#000000",
      }
    );

    const rectangle = scope.rect({
      x: 0,
      y: 150,
      width: 16,
      height: 200,
      alpha: 0.1,
      line_color: "red",
      line_alpha: 1,
    });

    const arrow = new Bokeh.Arrow({
      x_start: 0,
      y_start: 150,
      x_end: 0,
      y_end: 0,
      line_color: "red",
      line_alpha: 1,
      line_width: 3,
      line_cap: "round",
      end: new Bokeh.NormalHead({
        line_color: "red",
        fill_color: "red",
        size: 15,
      }),
    });

    scope_inset.add_layout(arrow);

    const doc = new Bokeh.Document();
    doc.add_root(scope);
    doc.add_root(scope_inset);

    Bokeh.embed.add_document_standalone(doc, document.getElementById("plot"));
    // Bokeh.embed.add_document_standalone(doc_inset, document.getElementById("plot-inset"));

    // slider_1 = document.getElementById("myRange_1");
    // slider_2 = document.getElementById("myRange_2");
    // output_1 = document.getElementById("value_1");
    // output_2 = document.getElementById("value_2");
    // output_1.innerHTML = slider_1.value; // Display the default slider value
    // output_2.innerHTML = slider_2.value; // Display the default slider value

    update_graph(waveformData, 0, 0);

    setTimeout(() => {
      let figures = document.querySelectorAll(".bk-Figure");
      let rel = figures[0];
      let abs = figures[1];
      let targetRect = rel.getBoundingClientRect();

      abs.style.top = targetRect.top + "px";
      abs.style.left = targetRect.left + (2 * targetRect.width) / 3 + "px";
      abs.style.width = targetRect.width / 3 + "px";
      abs.style.height = targetRect.height / 2 + "px";
    }, 20);
    console.log("graph made");

    // // Update the current slider value (each time you drag the slider handle)
    // slider_1.oninput = handle_sliders;
    // slider_2.oninput = handle_sliders;
  }

  if (browser) {
    onMount(async () => {
      await setup();
    });
  }
  // setup()
  // slider_value_1.set(500)
  // slider_value_2.set(500)

  // slider_value_1.set(1000)
  // slider_value_2.set(1500)
</script>

<div class="container">
  <div id="plot" />
  <div id="plot-inset" />

  <!-- {#await setup()}
    <p>waiting</p>
    {:then stuff} -->
  <div class="slidecontainer">
    <input
      type="range"
      min="1"
      max="2000"
      bind:value={$slider_value_1}
      class="slider"
      id="myRange_1"
    />
  </div>
  <p class="output">{Math.round($slider_value_1)}</p>
  <!-- <div id="value_1" class="output"/> -->

  <div class="slidecontainer">
    <input
      type="range"
      min="1"
      max="2000"
      bind:value={$slider_value_2}
      class="slider"
      id="myRange_2"
    />
  </div>
  <p class="output">{Math.round($slider_value_2)}</p>
  <!-- <div id="value_2" class="output"/> -->
  <!-- {/await} -->
</div>

<!-- <button on:click={() => slider_value_1.set(0)}>
      0%
  </button>
  
  <button on:click={() => slider_value_1.set(400)}>
      25%
  </button>
  
  <button on:click={() => slider_value_1.set(800)}>
      50%
  </button> -->

<style>
  #plot-inset {
    position: absolute;
    top: 10%;
    left: 80%;
    z-index: 9999;
    /* width: 700px; */
  }

  /* select the second instance of class bk-figure */
  :global(#plot div:nth-of-type(2)) {
    position: absolute;
    top: 10%;
    left: 80%;
    z-index: 9999;
    /* width: 700px; */
  }

  .container {
    margin: 5rem 20rem;
    /* margin: auto; */
    /* max-width: 60rem; */
  }
  .slidecontainer {
    width: 100%; /* Width of the outside container */
    padding-top: 5rem;
  }

  /* The slider itself */
  .slider {
    -webkit-appearance: none; /* Override default CSS styles */
    appearance: none;
    width: 100%; /* Full-width */
    height: 25px; /* Specified height */
    background: #e5e5e5; /* Grey background */
    outline: none; /* Remove outline */
    opacity: 0.7; /* Set transparency (for mouse-over effects on hover) */
    -webkit-transition: 0.2s; /* 0.2 seconds transition on hover */
    transition: opacity 0.2s;
    border-radius: 0.3rem;
  }

  /* Mouse-over effects */
  .slider:hover {
    opacity: 1; /* Fully shown on mouse-over */
  }

  /* The slider handle (use -webkit- (Chrome, Opera, Safari, Edge) and -moz- (Firefox) to override default look) */
  .slider::-webkit-slider-thumb {
    -webkit-appearance: none; /* Override default look */
    appearance: none;
    width: 25px; /* Set a specific slider handle width */
    height: 25px; /* Slider handle height */
    background: #9a9a9a; /* Green background */
    cursor: pointer; /* Cursor on hover */
    border-radius: 0.3rem;
  }

  .slider::-moz-range-thumb {
    width: 25px; /* Set a specific slider handle width */
    height: 25px; /* Slider handle height */
    background: #04aa6d; /* Green background */
    cursor: pointer; /* Cursor on hover */
  }

  .output {
    /* font style */
    font-family: "Roboto", sans-serif;
    font-weight: bold;
  }
</style>
