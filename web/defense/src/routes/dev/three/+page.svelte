<script>
    import * as THREE from "three";
    import { OrbitControls } from "three/examples/jsm/controls/OrbitControls.js";
    import * as d3 from "d3";
    import { onMount } from "svelte";
    import { browser } from "$app/environment";

    function f(x, y, t) {
        // inward ripple wave
        let phase1 =
            (0.5 +
                0.5 *
                    Math.sin(
                        t / 15 + 0 * Math.PI + 0.5 * Math.sqrt(x ** 2 + y ** 2)
                    )) **
            2;
        let rippleWave = Math.cos(
            2 * t + 25 * Math.sqrt(0.2 + x ** 2 + y ** 2)
        );

        // outward spiral wave
        let phase2 =
            (0.5 +
                0.5 *
                    Math.sin(
                        t / 15 +
                            0.75 * Math.PI -
                            0.5 * Math.sqrt(x ** 2 + y ** 2)
                    )) **
            2;
        let spiralWave =
            (Math.cos(
                // wave shape
                -2 * t + // animate
                    18 * Math.sqrt(0.2 + x ** 2 + y ** 2) - // outward ripple
                    2 * Math.atan((x + 0.001) / y) // spiral ramp
            ) *
                (x ** 2 + y ** 2)) /
                (x ** 2 + y ** 2 + 0.03) + // smooth roughness near origin
            1 -
            (x ** 2 + y ** 2) / (x ** 2 + y ** 2 + 0.03); // adds a bump so the spiral arms connect

        // chaotic inverse wave
        let phase3 = Math.max(0, 1 - phase1 - phase2);
        let inverseWave = Math.sin(
            (20 * (x * y)) / Math.sqrt(1 + x ** 2 + y ** 2) - t
        );
        // return (x, y, t)
        return (
            0.13 * phase1 * rippleWave +
            0.15 * phase2 * spiralWave +
            0.2 * phase3 * inverseWave +
            0.35
        );
    }

    function transform(x, y, z, t) {
        //transform matrix
        let tx = 1 * x + 0 * y + 0 * z;
        let ty = 0 * x + 1 * y + 0 * z;
        let tz = 0 * x + 0 * y + 1 * z;
        return [tx, ty, tz];
    }


    // function _domain(){return(
    // {  x:[-1,1],  y:[-1,1],  z:[-1,1], ticks:5}
    // )}

    function _graph(width, options, domain, f, transform) {
        // SET UP ------------------------------------
        // scene
        const aspect = 0.4;
        //const width = 600
        const height = width * aspect;
        const scene = new THREE.Scene();
        const pi = Math.PI;

        // MATERIALS ---------------------------------
        const matsurface = options.surface;
        const matSolidLine = options.line;

        // GRAPH ------------------------------------
        //const axesHelper = new THREE.AxesHelper( 5 );
        //scene.add( axesHelper )

        var centerHelper = new THREE.Group();
        var gridHelper = new THREE.Group();

        let xScale = d3.scaleLinear([0, 1], domain.x);
        let yScale = d3.scaleLinear([0, 1], domain.y);

        // console.log(xScale(0.5), yScale(0.5)

        let xRes = 10 * 10 + 1;
        let yRes = 10 * 10 + 1;

        let indices = [];

        // console.log("this: ", yScale(70 / (yRes - 1)))
        let diagonal;
        let centerPoint = [0, 0, 0];
        for (let i = 0; i < xRes - 1; i++) {
            for (let j = 0; j < yRes - 1; j++) {
                diagonal =
                    true ==
                    ((xScale((i + 1) / xRes) < centerPoint[0] &&
                        yScale((j + 1) / yRes) < centerPoint[1]) ||
                        (xScale((i + 1) / xRes) > centerPoint[0] &&
                            yScale((j + 1) / yRes) > centerPoint[1]));
                if (diagonal) {
                    indices.push(
                        i + j * xRes,
                        i + 1 + j * xRes,
                        i + (j + 1) * xRes,
                        i + 1 + j * xRes,
                        i + 1 + (j + 1) * xRes,
                        i + (j + 1) * xRes
                    );
                } else {
                    indices.push(
                        i + j * xRes,
                        i + 1 + j * xRes,
                        i + 1 + (j + 1) * xRes,
                        i + j * xRes,
                        i + 1 + (j + 1) * xRes,
                        i + (j + 1) * xRes
                    );
                }
            }
        }

        let vertices = [];
        xRes = 10 * 10 + 1;
        yRes = 10 * 10 + 1;
        for (let i = 0; i < yRes; i++) {
            for (let j = 0; j < xRes; j++) {
                vertices.push(
                    yScale(i / (yRes - 1)),
                    0.01,
                    xScale(j / (xRes - 1))
                );
            }
        }

        // vertices = new Float32Array(vertices);
        console.log("vertices new: ", vertices);

        // After the loop that generates the vertices
        for (let i = 0; i < vertices.length; i++) {
            if (isNaN(vertices[i])) {
                console.error(`vertices[${i}] is NaN`);
            }
        }
        console.log("vertices length", vertices.length);

        let surfaceGeo = new THREE.BufferGeometry();
        surfaceGeo.setAttribute(
            "position",
            new THREE.Float32BufferAttribute(vertices, 3)
        );

        console.log("this is the verticles: ", vertices);
        console.log(
            "this is the position: ",
            surfaceGeo.attributes.position.array
        );
        // console.log("count? ", surfaceGeo.count)
        // console.log("set attribute")
        surfaceGeo.setIndex(indices);
        let surfaceMesh = new THREE.Mesh(surfaceGeo, matsurface);

        // console.log("surface mesh is", surfaceMesh)

        // console.log("surface geo is", surfaceGeo);
        // console.log(new THREE.Mesh(surfaceGeo, matsurface))

        //wireframe
        let gridList = [];
        let ticks = domain.ticks;
        for (let i = 0; i < ticks; i++) {
            let xVerts = [];
            let yVerts = [];
            const xGeo = new THREE.BufferGeometry();
            const yGeo = new THREE.BufferGeometry();
            for (let j = 0; j < yRes; j++) {
                xVerts.push(yScale(j / (yRes - 1)), 0, xScale(i / (ticks - 1)));
            }
            for (let j = 0; j < xRes; j++) {
                yVerts.push(yScale(i / (ticks - 1)), 0, xScale(j / (xRes - 1)));
            }

            for (let i = 0; i < xVerts.length; i++) {
                if (isNaN(xVerts[i])) {
                    console.error(`vertices[${i}] is NaN`);
                }
            }
            xGeo.setAttribute(
                "position",
                new THREE.Float32BufferAttribute(xVerts, 3)
            );
            yGeo.setAttribute(
                "position",
                new THREE.Float32BufferAttribute(yVerts, 3)
            );
            const xLine = new THREE.Line(xGeo, matSolidLine);
            const yLine = new THREE.Line(yGeo, matSolidLine);
            gridList.push(xLine, yLine);
            gridHelper.add(xLine, yLine);
        }
        gridHelper.position.set(
            0,
            0.0005 * Math.abs(domain.z[0] - domain.z[1]),
            0
        );
        // console.log("wireframe setup");
        // console.log(Math.abs(domain.x[0] - domain.x[1]));

        centerHelper.add(surfaceMesh);
        centerHelper.add(gridHelper);
        scene.add(centerHelper);
        centerHelper.scale.set(
            2 / Math.abs(domain.y[0] - domain.y[1]),
            2 / Math.abs(domain.z[0] - domain.z[1]),
            2 / Math.abs(domain.x[0] - domain.x[1])
        );
        let xMove =
            -(domain.x[0] + domain.x[1]) / Math.abs(domain.x[0] - domain.x[1]);
        let yMove =
            -(domain.y[0] + domain.y[1]) / Math.abs(domain.y[0] - domain.y[1]);
        let zMove =
            -(domain.z[0] + domain.z[1]) / Math.abs(domain.z[0] - domain.z[1]);
        centerHelper.position.set(yMove, zMove, xMove);

        // camera
        const camera = new THREE.PerspectiveCamera(45, width / height, 1, 1000); // field of view, aspect ratio, near, far plane
        camera.position.set(-1.75, 1.25, 1.75);
        camera.lookAt(0, 0, 0);

        // renderer
        const renderer = new THREE.WebGLRenderer({ antialias: true }); // most powerful all purpose renderer, with smooth edges
        renderer.setClearColor("#ffffff"); // background color #f6f6f6
        renderer.setSize(width, height); // set window size

        // camera controls
        const controls = new OrbitControls(camera, renderer.domElement); // allows mouse to move camera
        controls.enableZoom = false; // disables zoom
        controls.enablePan = false; // disables pan
        controls.enableRotate = true; // determines if camera can rotate

        // Light ---------------------------------------
        const hemiLight = new THREE.HemisphereLight(0xffffff, 0xffffff, 1);
        hemiLight.color.setHSL(0.6, 0.6, 0.5);
        hemiLight.groundColor.setHSL(0.095, 0.5, 0.75);
        hemiLight.position.set(0, 50, 0);
        scene.add(hemiLight);

        // console.log(hemiLight.position);

        const dirLight1 = new THREE.DirectionalLight(0xffffff, 1);
        dirLight1.color.setHSL(0.8 / 6, 0.7, 0.5);
        dirLight1.position.set(
            Math.sin((1.5 / 3) * pi),
            2,
            Math.cos((2 / 3) * pi)
        );
        dirLight1.position.multiplyScalar(50);
        scene.add(dirLight1);

        const dirLight2 = new THREE.DirectionalLight(0xffffff, 1);
        dirLight2.color.setHSL(3 / 6, 0.2, 0.5);
        dirLight2.position.set(
            Math.sin((3.5 / 3) * pi),
            1,
            Math.cos((4 / 3) * pi)
        );
        dirLight2.position.multiplyScalar(50);
        scene.add(dirLight2);

        const dirLight3 = new THREE.DirectionalLight(0xffffff, 0.5);
        dirLight3.color.setHSL(5 / 6, 0.4, 0.5);
        dirLight3.position.set(
            Math.sin((5.5 / 3) * pi),
            1,
            Math.cos((0 / 3) * pi)
        );
        dirLight3.position.multiplyScalar(50);
        scene.add(dirLight3);

        //////////////////////////////////////////////////////////////////////////////////////////////
        // RENDER LOOP ///////////////////////////////////////////////////////////////////////////////
        //////////////////////////////////////////////////////////////////////////////////////////////
        let t = 0;
        let xpos, ypos, zpos, transformList, savedPoints;

        function savePoints(mesh) {
            const positions = mesh.geometry.attributes.position.array;
            mesh.savedPoints = [];
            for (let i = 0; i < positions.length; i += 3) {
                mesh.savedPoints.push({ x: positions[i], z: positions[i + 2] });
            }
        }
        savePoints(surfaceMesh);
        for (let i = 0; i < gridList.length; i++) {
            savePoints(gridList[i]);
        }
        let thing = 3;

        /**
         * @param {THREE.Mesh} mesh
         */
        function updateSurface(mesh, t) {
            const positions = mesh.geometry.attributes.position.array;
            const count = positions.length / 3;
            for (let i = 0; i < count; i++) {
                const xpos = mesh.savedPoints[i].x;
                // console.log(xpos)
                const ypos = mesh.savedPoints[i].z;
                const zpos = f(xpos, ypos, t);
                // console.log("the nan thing: ", zpos);
                const transformList = transform(xpos, ypos, zpos, t);
                // console.log(transformList)
                if (thing < 10) {
                    // console.log("transform list is", transformList);
                    thing += 1;
                }
                positions[i * 3] = transformList[0];
                positions[i * 3 + 2] = transformList[1];
                positions[i * 3 + 1] = transformList[2];
            }
            mesh.geometry.attributes.position.needsUpdate = true;
            mesh.geometry.computeVertexNormals();
        }

        renderer.domElement.update = function render() {
            // console.log("grid list is", gridList);
            t = new Date().getTime() / 1000; // time in seconds

            updateSurface(surfaceMesh, t);
            for (let i = 0; i < gridList.length; i++) {
                updateSurface(gridList[i], t);
            }
            renderer.render(scene, camera); // renders scene
            // console.log("rendering");

            return "update render externally";
        };

        // Input -------------------------
        return renderer.domElement;
    }

    let width = 0;

    const domain = { x: [-1, 1], y: [-1, 1], z: [-1, 1], ticks: 5 };
    const options = {surface: new THREE.MeshStandardMaterial({
                color: "lightblue",
                roughness: 0.5,
                metalness: 0.02,
                side: THREE.DoubleSide,
                wireframe: false,
            }),
            line: new THREE.LineBasicMaterial({
                color: "blue",
                transparent: true,
                opacity: 0.15,
                // depthwrite: false,
            }),
        };

    // _graph(width,_options,domain,_f,_transform)

    let container;
    let canvas;

    if (browser) {
        onMount(() => {
            if (typeof window !== "undefined") {
                width = window.innerWidth;
            }
            // container = document.getElementById("my-container");
            canvas = _graph(width, options, domain, f, transform);
            container.appendChild(canvas);

            function animate() {
                canvas.update();
                requestAnimationFrame(animate);
            }

            animate();
        });
    }
</script>

<h1>this is my content above</h1>

<div bind:this={container} />
<h1>this is my content below</h1>

<!-- <script>
    import { onMount } from 'svelte';
    import * as THREE from 'three';
  
    let container;
  
    onMount(() => {
      const scene = new THREE.Scene();
      const camera = new THREE.PerspectiveCamera(75, window.innerWidth / window.innerHeight, 0.1, 1000);
      const renderer = new THREE.WebGLRenderer();
  
      container.appendChild(renderer.domElement);
  
      const geometry = new THREE.BoxGeometry();
      const material = new THREE.MeshBasicMaterial({ color: 0x00ff00 });
      const cube = new THREE.Mesh(geometry, material);
      scene.add(cube);
  
      camera.position.z = 5;
  
      const animate = function () {
        requestAnimationFrame(animate);
  
        cube.rotation.x += 0.01;
        cube.rotation.y += 0.01;
  
        renderer.render(scene, camera);
      };
  
      animate();
    });
  </script>
  
  <div bind:this={container}></div> -->
