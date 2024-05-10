const xlsx = require('xlsx');
const fs = require('fs');
const path = require('path');

// Function to print usage instructions
const printHelp = () => {
  console.log('Usage: node index.js <input.xls> [<output.xlsx>]');
};

// Function to parse the output path
const parseOutput = (input, output) => {
  let outDir;
  const i = path.parse(input);
  if (output) {
    const o = path.parse(output);
    if (o.ext === '.xlsx') {
      return output;
    }
    outDir = output;
  } else {
    outDir = i.dir;
  }
  let filename = path.join(outDir, `${i.name}.xlsx`);
  if (fs.existsSync(filename)) {
    filename = path.join(outDir, `${i.name}-${new Date().toJSON().replace(/[-:]/g, '')}.xlsx`);
  }
  return filename;
};

// Function to perform the conversion
const convert = (input, output) => {
  // Read the input .xls file
  const workbook = xlsx.readFile(input);
  // Parse the output path
  output = parseOutput(input, output);
  // Write the converted .xlsx file
  xlsx.writeFile(workbook, output);
  console.log(`Converted '${input}' to '${output}'`);
};

// Main function to parse command line arguments and start the conversion process
const main = () => {
  const input = 'AMMASANDRA_110_15_10_2023.xls'; // Specify the input file
  const output = process.argv[2]; // Get the output file path from command line argument
  // Check if input file is a .xls file
  if (!input.endsWith('.xls')) {
    console.error('Input file is not a .xls file');
    process.exit(1);
  }
  // Check if input file exists
  if (!fs.existsSync(input)) {
    console.error(`File does not exist: ${input}`);
    process.exit(1);
  }

  // Perform the conversion
  convert(input, output);
};

// Call the main function when the script is executed directly
if (require.main === module) {
  main(); // Call the main function
}


