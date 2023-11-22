<script lang="ts">
    import { random } from "lodash";

    let rough_amplitudes_x: Array<number> = [];
    let rough_amplitudes_y: Array<number> = [];

    let speed = 0.1;
    let initial_number = 20;
    for (let amp_idx = 0; amp_idx < initial_number; amp_idx++) {
        let noise = random(-0.1, 0.1, true);
        let sinValue = Math.sin(amp_idx * speed);
        rough_amplitudes_x.push(amp_idx / initial_number);
        rough_amplitudes_y.push(sinValue + noise);
    }
    let amp_idx = initial_number;

    // there are several parameters for graphing the pll line that
    // intermittently get updated

    // I dont want the graph range update to be jerky. Does this mean I should be using graph limits that get updated?
    // or... just the data moves. The graph is stationary.
    // so we always add x = 1 to rough_amplitudes_x and

    // animate function inside requestAnimationFrame
    function animate() {
        let start = performance.now();
        let rough_amplitude_add_time = start;

        requestAnimationFrame(function animate(time) {
            // timeFraction goes from 0 to 1
            // if 1 second has passed at a point to the array rough_amplitudes
            if (time - rough_amplitude_add_time > 1000) {
                rough_amplitudes_x.push(1);
                // compute the sin of rought_amplitude_idx and add some random noise to it
                let noise = random(-0.1, 0.1, true); // generate a random float between -1 and 1
                let sinValue = Math.sin(amp_idx * speed);
                rough_amplitudes_y.push(sinValue + noise);
                rough_amplitude_add_time = time;
                amp_idx++;
            }

            // // for value in rough_amplitudes x...
            // rough_amplitudes_x.forEach((x, idx) => {
            //     // if the value is greater than 1, remove it from the array
            //     x -= 0.01;

            //     if (x < 0) {
            //         rough_amplitudes_x.splice(idx, 1);
            //         rough_amplitudes_y.splice(idx, 1);
            //     }
            // });

            // this is interesting...
            let paired = rough_amplitudes_x.map((x, idx) => [x - 0.01, rough_amplitudes_y[idx]]);
            let filtered = paired.filter(pair => pair[0] < 0);

            rough_amplitudes_x = filtered.map(pair => pair[0]);
            rough_amplitudes_y = filtered.map(pair => pair[1]);
           



            requestAnimationFrame(animate);
        });
    }
</script>
