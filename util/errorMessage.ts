async function errorMessage() {
	const sleep = (ms: number) => new Promise(r => setTimeout(r, ms));

	console.log("TESTING TESTING...")
	await sleep(1000)
	for (let i = 0; i < 3; i++) {
		console.log("SCANNING...")
		await sleep(1000)
	}
	for (let i = 0; i < 10; i++) {
		console.log("🚨🚨🚨 DUMBASS DETECTED 🚨🚨🚨")
		await sleep(500)
	}
	console.log("ERROR 69420: excess of dumbassery")
	await sleep(500)
	console.log("ABORTING PROGRAM")
}

export default errorMessage
