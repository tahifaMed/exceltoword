package com.cheuvreux.copernic.cli;

import org.apache.commons.cli.CommandLine;
import org.apache.commons.cli.HelpFormatter;
import org.apache.commons.cli.Option;
import org.apache.commons.cli.Options;

public class CLIParameters {
    public static Options configParameters(final Options firstOptions) {

        final Option excelFile = Option.builder("x")
                .longOpt("excel")
                .desc("Excel File to read source and destination from")
                .hasArg(true)
                .required(true)
                .build();

        final Option sheetIndex = Option.builder("si")
                .longOpt("sheet-index")
                .desc("the sheet index in the excel File, default 0")
                .hasArg(true)
                .optionalArg(true)
                .required(false)
                .build();

        final Option columnSourceIndex = Option.builder("csi")
                .longOpt("column-source-index")
                .desc("the row index of the source column, default 2")
                .hasArg(true)
                .required(false)
                .optionalArg(true)
                .build();

        final Option columnDestinationIndex = Option.builder("cdi")
                .longOpt("column-destination-index")
                .desc("the row index of the source column default 3")
                .hasArg(true)
                .required(false)
                .optionalArg(true)
                .build();

        final Option wordInput = Option.builder("w")
                .longOpt("word")
                .desc("Word File to read From")
                .hasArg(true)
                .required(true)
                .build();

        final Options options = new Options();
        // First Options
        for (final Option fo : firstOptions.getOptions()) {
            options.addOption(fo);
        }

        options.addOption(excelFile);
        options.addOption(wordInput);
        options.addOption(sheetIndex);
        options.addOption(columnSourceIndex);
        options.addOption(columnDestinationIndex);

        return options;
    }

    public static Options configFirstParameters() {

        final Option helpFileOption = Option.builder("h")
                .longOpt("help")
                .desc("show help message")
                .build();

        final Options firstOptions = new Options();

        firstOptions.addOption(helpFileOption);

        return firstOptions;
    }

    public static void helpMode(Options options, CommandLine firstLine) {
        // Si mode aide
        boolean helpMode = firstLine.hasOption("help");
        if (helpMode) {
            final HelpFormatter formatter = new HelpFormatter();
            formatter.printHelp("Extract From Excel And replace in Word", options, true);
            System.exit(0);
        }
    }
}
