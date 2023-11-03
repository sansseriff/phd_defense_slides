<script>
	import Counter from "./Counter.svelte";
	import welcome from "$lib/images/svelte-welcome.webp";
	import welcome_fallback from "$lib/images/svelte-welcome.png";
	import Header from "./Header.svelte";
</script>

<svelte:head>
	<title>Home</title>
	<meta name="description" content="Svelte demo app" />
</svelte:head>

<section>
	<!-- <h1>
		<span class="welcome">
			<picture>
				<source srcset={welcome} type="image/webp" />
				<img src={welcome_fallback} alt="Welcome" />
			</picture>
		</span>

		SNSPDs for Fun and Profit
	</h1> -->
	<h1 class="gradient-text">SNSPDs for Fun & Profit</h1>
	<h3>Andrew Mueller</h3>
	<h3>9/6/2023</h3>

	<!-- <h2>
		try editing <strong>src/routes/+page.svelte</strong>
	</h2> -->

	<!-- <Counter /> -->
</section>

<style>
	.gradient-text {
		background: linear-gradient(to right, #ff00ff, #00ffff);
		-webkit-background-clip: text;
		background-clip: text;
		-webkit-text-fill-color: transparent;
	}
	section {
		display: flex;
		flex-direction: column;
		min-height: 80vh;
		display: flex;
		flex-direction: column;
		justify-content: center;
		align-items: center;
		flex: 0.6;
	}

	h1 {
		width: 100%;
	}

	/* .welcome {
		display: block;
		position: relative;
		width: 100%;
		height: 0;
		padding: 0 0 calc(100% * 495 / 2048) 0;
	} */

	/* .welcome img {
		position: absolute;
		width: 100%;
		height: 100%;
		top: 0;
		display: block;
	} */
</style>
