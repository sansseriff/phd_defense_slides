<script lang="ts">
	import { page } from "$app/stores";
	import logo from "$lib/images/svelte-logo.svg";
	import github from "$lib/images/github.svg";
	import { onMount } from "svelte";
	import { writable } from "svelte/store";
	// import { connectToBackend } from '../lib/ws_control'
	

	const state = writable({
		isHeaderVisible: false,
	});

	onMount(async () => {

		

		
		// const ws_lib = await import("../lib/ws_control");
		// const result = ws_lib.connectToBackend();
		// // const slide_state: any = null;

		// // if (!result) {
		// // 	console.error("Failed to connect to backend");
		// // } else {
		// // 	const { connection, slide_state } = result;

		// // 	// Now you can use connection and slide_state safely
		// // 	// ...
		// // }
		// const { connection, slide_state } = result;

		// console.log("running?")
		// console.log(slide_state);

		// slide_state.slide += 1;

		// console.log(slide_state)
		// slide_state.slide += 1;
		// console.log(slide_state)

		const header = document.querySelector("header");

		document.addEventListener("mousemove", (event) => {
			const distanceFromTop = Math.abs(event.clientY - header.offsetTop);
			if (distanceFromTop < 50) {
				state.set({ isHeaderVisible: true });
			} else {
				state.set({ isHeaderVisible: false });
			}
		});
	});
</script>

<header class:visible={$state.isHeaderVisible}>
	<div class="corner">
		<!-- <a href="https://kit.svelte.dev">
			<img src={logo} alt="SvelteKit" />
		</a> -->
	</div>

	<nav>
		<svg viewBox="0 0 2 3" aria-hidden="true">
			<path
				d="M0,0 L1,2 C1.5,3 1.5,3 2,3 L2,0 Z"
				style="fill: #F0F0F0;"
			/>
		</svg>
		<ul>
			<li aria-current={$page.url.pathname === "/" ? "page" : undefined}>
				<a href="/">Home</a>
			</li>
			<!-- <li aria-current={$page.url.pathname === '/about' ? 'page' : undefined}>
				<a href="/about">About</a>
			</li> -->
			<!-- <li aria-current={$page.url.pathname.startsWith('/sverdle') ? 'page' : undefined}>
				<a href="/sverdle">Sverdle</a>
			</li> -->
			<!-- <li aria-current={$page.url.pathname.startsWith('/my_page') ? 'page' : undefined}>
				<a href="/my_page">my page</a>
			</li> -->
			<li
				aria-current={$page.url.pathname.startsWith("/slide_0")
					? "page"
					: undefined}
			>
				<a href="/slide_0">slides</a>
			</li>
		</ul>
		<svg viewBox="0 0 2 3" aria-hidden="true">
			<path
				d="M0,0 L0,3 C0.5,3 0.5,3 1,2 L2,0 Z"
				style="fill: #F0F0F0;"
			/>
		</svg>
	</nav>

	<div class="corner">
		<a href="https://github.com/sveltejs/kit">
			<img src={github} alt="GitHub" />
		</a>
	</div>
</header>

<style>
	header {
		display: flex;
		justify-content: space-between;
		opacity: 0;
		transition: opacity 0.2s ease-in-out;
		position: absolute;
		top: 0;
		left: 0;
		width: 100%;
		z-index: 9999;
		opacity: 0;
		transition: opacity 0.2s ease-in-out;
		/* background-color: red; */
	}

	header.visible {
		opacity: 1;
		/* background-color: green; */
	}

	.corner {
		width: 3em;
		height: 3em;
	}

	.corner a {
		/* display: flex; */
		align-items: center;
		justify-content: center;
		width: 100%;
		height: 100%;
	}

	.corner img {
		width: 2em;
		height: 2em;
		object-fit: contain;
	}

	nav {
		display: flex;
		justify-content: center;
		--background: rgba(255, 255, 255, 0.7);
	}

	svg {
		width: 2em;
		height: 3em;
		display: block;
	}

	path {
		fill: var(--background);
	}

	ul {
		position: relative;
		padding: 0;
		margin: 0;
		height: 3em;
		display: flex;
		justify-content: center;
		align-items: center;
		list-style: none;
		background: var(--background);
		background-size: contain;
	}

	li {
		position: relative;
		height: 100%;
		background-color: #f0f0f0;
	}

	li[aria-current="page"]::before {
		--size: 6px;
		content: "";
		width: 0;
		height: 0;
		position: absolute;
		top: 0;
		left: calc(50% - var(--size));
		border: var(--size) solid transparent;
		border-top: var(--size) solid var(--color-theme-1);
	}

	nav a {
		display: flex;
		height: 100%;
		align-items: center;
		padding: 0 0.5rem;
		color: var(--color-text);
		font-weight: 700;
		font-size: 0.8rem;
		text-transform: uppercase;
		letter-spacing: 0.1em;
		text-decoration: none;
		transition: color 0.2s linear;
	}

	a:hover {
		color: var(--color-theme-1);
	}
</style>
