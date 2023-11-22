<script lang="ts">
    // import svelteLogo from "./assets/svelte.svg";
    // import viteLogo from "/vite.svg";
    // import Counter from "./lib/Counter.svelte";
    import { onMount } from "svelte";
    import { randomNormal } from "d3-random";
    import { diff, i, re, round } from "mathjs";
    import { tweened } from "svelte/motion";
    import { cubicOut } from "svelte/easing";
    import { tick } from "svelte";
    import { browser } from "$app/environment";

    interface JsonData {
        x: number[];
        y: number[];
        sxs: number[];
        t1s: number[];
        x_points: number[];
        y_points: number[];
    }

    let scope = null;
    let points = null;

    let data: any;

    async function setup() {
        const jsonData = await loadJsonData("/pulse_distort_data.json");
        make_graph(jsonData);
    }

    async function loadJsonData(path: string): Promise<JsonData> {
        const response = await fetch(path);
        return await response.json();
    }

    function transform_x_cut(x_array: Array<number>, current_sxs: number, current_t1s: number): Array<number> {
        return x_array.map((x, idx) => {
            return (3*x + 1*x * current_sxs)/4 + 58 + 0.5 * current_t1s;
        });
    }

    function transform_y_cut(y_array: Array<number>, current_x_pt: number, current_y_pt: number): Array<number> {
        return y_array.map((y, idx) => {
            return (y + y*500/(current_y_pt - current_x_pt))/2
        });
    }

    async function make_graph(d: JsonData) {
        // waveformData = {
        //     sliding_pulse_1: json_data.y.slice(),
        //     sliding_pulse_2: json_data.y.slice(),
        //     existing_waveform: json_data.y.slice(),
        //     current_waveform: json_data.y.slice(),
        // };

        // let pulse_offset = createRandomArray(1300, 0.2);

        // waveformData.existing_waveform.splice(
        //     -pulse_offset.length,
        //     pulse_offset.length
        // );
        // waveformData.existing_waveform.unshift(...pulse_offset);

        // data = new Bokeh.ColumnDataSource({
        //     data: {
        //         x: json_data.x.map((value) => value - 3200 + 84),
        //         y: waveformData.current_waveform,
        //         y_stationary: waveformData.existing_waveform,
        //         y_sliding_1: waveformData.sliding_pulse_1,
        //         y_sliding_2: waveformData.sliding_pulse_2,
        //     },
        // });

        // color: [black,
        //         grey,
        //         grey,  
        //         grey,
        //         ...]
        let array_of_arrays_x: Array<Array<number>> = []
        let array_of_arrays_y: Array<Array<number>> = []
        
        
        

        d.x.forEach((value, idx) => {
            array_of_arrays_x.push(transform_x_cut(d.x_points, d.sxs[idx], d.t1s[idx]))
            array_of_arrays_y.push(transform_y_cut(d.y_points, d.x[idx], d.y[idx]))
        });

        array_of_arrays_x.length = 100
        array_of_arrays_y.length = 100
        let zeros = new Array(array_of_arrays_x.length).fill(0.1);
        // zeros[3] = 0.4;
        // zeros[0] = 1;
        // zeros[1] = 0.8;
        // zeros[2] = 0.6;
        
        // zeros[4] = 0.2;
        // zeros[5] = 0;

        zeros.length = 100

        // console.log(zeros)

        let special = new Bokeh.ColumnDataSource({
            data: {
                x_arrrr: array_of_arrays_x,
                y_arrrr: array_of_arrays_y,
                alphas: zeros,
            },
        });

        console.log(array_of_arrays_x[0].length)

        
        data = new Bokeh.ColumnDataSource({
            data: {
                // growing_x: [],
                // growing_y: [],
                // x: d.x,
                // y: d.y,
                // sxs: d.sxs,
                // t1s: d.t1s,
                x_points: d.x_points,
                y_points: d.y_points,
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
            sizing_mode: "stretch_height",
            height: 600,
            width: 700,
            x_range: [0, 1000],
            y_range: [0, 70],
        });
        scope.toolbar.logo = null;

        points = new Bokeh.Plotting.figure({
            // title: "Example of random data",
            // tools: "pan,wheel_zoom,box_zoom,reset,save",
            tools: "",
            // toolbar: {logo: null},
            //tools: "",
            // sizing_mode: "stretch_both",
            sizing_mode: "stretch_height",
            height: 900,
            width: 700,
            x_range: [0, 900],
            y_range: [0, 60],
        });
        points.toolbar.logo = null;


        
        // const line_1 = scope.line(
        //     { field: "x_points" },
        //     { field: "y_points" },
        //     {
        //         source: data,
        //         line_width: 4,
        //         line_color: "#000000",
        //     }
        // );

        const lines_1 = scope.multi_line(
            { field: "x_arrrr" },
            { field: "y_arrrr" },
            {
                source: special,
                line_width: 4,
                line_color: "#000000",
                line_alpha: { field: "alphas"},
            }
        );

        // for (let j = 0; j < 30; j++) {
            
        // }

        // scope.add_layout(band_1);
        // scope.add_layout(band_2);

        const doc_scope = new Bokeh.Document();
        const doc_points = new Bokeh.Document();

        doc_scope.add_root(scope);
        doc_points.add_root(points);
        Bokeh.embed.add_document_standalone(
            doc_scope,
            document.getElementById("plot_scope")
        );

        Bokeh.embed.add_document_standalone(
            doc_points,
            document.getElementById("plot_points")
        );
        // Bokeh.embed.add_document_standalone(doc_inset, document.getElementById("plot-inset"));

        // slider_1 = document.getElementById("myRange_1");
        // slider_2 = document.getElementById("myRange_2");
        // output_1 = document.getElementById("value_1");
        // output_2 = document.getElementById("value_2");
        // output_1.innerHTML = slider_1.value; // Display the default slider value
        // output_2.innerHTML = slider_2.value; // Display the default slider value

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
</script>

<div class="container">
    <div id="plot_scope" />
    <div id="plot_points" />

</div>

<style>

#plot_scope {
    margin: 3rem;
    height: 60%
}

#plot_points {
    margin: 3rem;
    height: 100%
}

.container {
    box-sizing: border-box;
    height: 100vh;
    /* margin: 20rem; */
    padding-left: 10rem;
    padding-right: 10rem;
    padding-top: 19rem;
    padding-bottom: 19rem;
    /* margin: auto; */
    /* max-width: 60rem; */
    display:flex;
    flex-direction:row;
    justify-content: center;
    align-items: center;
    
  }
</style>
