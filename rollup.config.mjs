import typescript from '@rollup/plugin-typescript';
import terser from '@rollup/plugin-terser';

const umdOutput = {
	name: "docx",
	file: 'dist/docx-preview.js',
	sourcemap: true,
	format: 'umd',
	globals: {
		jszip: 'JSZip'
	}
};

export default args => {
	const config = {
		input: 'src/docx-preview.ts',
		output: [umdOutput],
		plugins: [typescript()]
	}

	if (args.environment == 'BUILD:production')
		config.output = [umdOutput,
			{
				...umdOutput,
				file: 'dist/docx-preview.min.js',
				plugins: [terser()]
			},
			{
				file: 'dist/docx-preview.mjs',
				sourcemap: true,
				format: 'es',
			},
			{
				file: 'dist/docx-preview.min.mjs',
				sourcemap: true,
				format: 'es',
				plugins: [terser()]
			}];

	return config
};