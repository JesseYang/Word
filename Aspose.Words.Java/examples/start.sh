# Set SUN_JRE_1_7_HOME to the home folder of SUN JRE
# At least SUN JRE 1.7 is required to successfully run the application. It is due to the JavaFX dependency.
# Linux Example: export SUN_JRE_1_7_HOME=/usr/lib/jvm/jre1.7.0_15
# Mac Example: export SUN_JRE_1_7_HOME=/Library/Internet\ Plug-Ins/JavaAppletPlugin.plugin/Contents/Home
export SUN_JRE_1_7_HOME=/usr/lib/jvm/jre1.7.0_15
echo SUN_JRE_1_7_HOME = ${SUN_JRE_1_7_HOME}

# Run the launcher program
"${SUN_JRE_1_7_HOME}/bin/java" -cp "lib/*:.:Aspose.Examples.Launcher.jar" aspose.examples.launcher.AsposeExamplesLauncher

# Find the absolute path of the Examples directory
export EXAMPLES_DIR="$( cd "$( dirname "${BASH_SOURCE[0]}" )" && pwd )"
EXAMPLES_DIR=${EXAMPLES_DIR}/Examples
echo EXAMPLES_DIR = ${EXAMPLES_DIR}

# Change directory to App
cd App

# Execute the start.sh script inside the App
sh startapp.sh

cd ..

# Clean the variables used
SUN_JRE_1_7_HOME=
EXAMPLES_DIR=

echo Finished execution of Aspose Java Dashboard Application
