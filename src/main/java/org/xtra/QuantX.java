package org.xtra;

import java.io.*;
import java.time.LocalDate;
import java.time.LocalTime;
import java.util.*;
import org.apache.commons.cli.CommandLine;
import org.apache.commons.cli.CommandLineParser;
import org.apache.commons.cli.DefaultParser;
import org.apache.commons.cli.HelpFormatter;
import org.apache.commons.cli.Option;
import org.apache.commons.cli.Options;
import org.apache.commons.cli.ParseException;

public class QuantX {

    private final List<String> fnGuideFilenames;
    private String fullCatalogFilename;
    private String shoppingGuideFilename;
    private String outputFilename;

    public QuantX() {
        LocalDate date = LocalDate.now();
        LocalTime time = LocalTime.now();

        this.fnGuideFilenames = new LinkedList<>();
        this.outputFilename = String.format("output-%d%02d%02d%02d%02d%02d.xlsx",
            date.getYear(), date.getMonthValue(), date.getDayOfMonth(),
            time.getHour(), time.getMinute(), time.getSecond());
    }

    public void addFnGuideFilename(String filename) {
        this.fnGuideFilenames.add(filename);
    }

    public void setFullCatalogFilename(String filename) {
        this.fullCatalogFilename = filename;
    }

    public void setShoppingGuideFilename(String filename) {
        this.shoppingGuideFilename = filename;
    }

    public void setOutputFilename(String filename) {
        this.outputFilename = filename;
    }

    public void run() throws IOException {

        Map<String, Company> companies = new HashMap<>();

        FullCatalog fullCatalog = new FullCatalog(companies);
        if (fullCatalogFilename != null) {
            fullCatalog.load(fullCatalogFilename);
        }

        for(String fnGuideFilename : fnGuideFilenames) {
            new FnGuide(companies).load(fnGuideFilename);
        }

        if (shoppingGuideFilename != null)  {
            new ShoppingGuide(companies).load(shoppingGuideFilename);
        }

        fullCatalog.calculate();
        fullCatalog.saveToExcel(this.outputFilename);

    }

    public static void main(String [] args) throws IOException {
        Options options = new Options();

        options.addOption(Option.builder("fn").hasArgs().valueSeparator(',').argName("FN Guide").longOpt("fn-guide").build());
        options.addOption(Option.builder("sg").hasArg().argName("Shopping Guide").longOpt("shopping-guide").build());
        options.addOption(Option.builder("fc").hasArg().argName("Full Catalog").longOpt("full-catalog").build());
        options.addOption(Option.builder("o").hasArg().argName("Output").longOpt("output").build());

        try {
            CommandLineParser parser = new DefaultParser();
            CommandLine commandLine = parser.parse(options, args);

            QuantX quantX = new QuantX();
            if (commandLine.hasOption("fc")) {
                quantX.setFullCatalogFilename(commandLine.getOptionValue("fc"));
            }

            if (commandLine.hasOption("fn")) {
                for(String filename : commandLine.getOptionValues("fn")) {
                    quantX.addFnGuideFilename(filename);
                }
            }

            if (commandLine.hasOption("sg")) {
                quantX.setShoppingGuideFilename(commandLine.getOptionValue("sg"));
            }

            if (commandLine.hasOption("o")) {
                quantX.setOutputFilename(commandLine.getOptionValue("o"));
            }

            quantX.run();
        } catch (ParseException e) {
            HelpFormatter helpFormatter = new HelpFormatter();
            helpFormatter.printHelp("QuantX", options);
        }
    }

}
