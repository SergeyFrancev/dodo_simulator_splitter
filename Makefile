build-release:
	# rm ./release/dodo_simulator_splitter__mac
	cargo build --release
	cp ./target/release/dodo_simulator_splitter ./release/dodo_simulator_splitter__mac

build-release-win:
	# rm ./release/dodo_simulator_splitter__win86.exe
	cargo build --target=x86_64-pc-windows-gnu --release
	cp ./target/x86_64-pc-windows-gnu/release/dodo_simulator_splitter.exe ./release/dodo_simulator_splitter__win86.exe