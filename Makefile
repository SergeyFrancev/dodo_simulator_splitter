build-release:
	rm ./target/release/dodo_simulator_splitter
	cargo build --release
	cp ./target/release/dodo_simulator_splitter ./test

build-release-win:
	cargo build --target=x86_64-pc-windows-gnu --release