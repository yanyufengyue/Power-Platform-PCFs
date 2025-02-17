Developer Information for developing Microsoft Power App Component
	1. Install NPM - https://nodejs.org/en/
	2. Powerapp CLI
	3. Visual studio 20xx

Commands
	1. pac pcf init --namespace <specify your namespace here> --name <put component name here> -- template <component type>
	
	2. The next step is to install project dependencies. Your component project would require lots of external files to build successfully.
	   To install the dependencies, run the command "npm install"
	
	3. Do coding in project directory created from command 1
	
	4. To build, you need to open Visual Studio 20xx developer command prompot.
	   Once in the directory, run the command "npm run build"

	5. Just type "npm start" in command line. It will launch Harness and show you.

	6. Now we need to create a solution file once all code part is done. Create a new folder inside the control folder, and go to the folder by:
	   cd foldername, then run the command below
	   pac solution init --publisher-name <enter your publisher name> --publisher-prefix <enter your publisher prefix>

	7. Next step to add our custom control reference to project with below command
	   We need in the deploy folder created in step 6.
	   pac solution add-reference --path <path or relative path of your PowerApps component framework project on disk>
	   The path is the directory of Parent folder of the PCF project, put the path with double qoute("C:\users\Lee Li\source\PCF") if there is a space in the folder path
	pac solution add-reference --path ..\ (if the solution folder is inside the project folder)
	
	8. Now we will generate solution file to import in D365 solution to be available for D365 with below command
 	   msbuild /t:restore
	   Once compleled without error then
	   msbuild
msbuild /t:build /restore (In the Solution folder)
msbuild /t:build /restore /p:configuration=Release
 
	   
General Commands:
Go back: cd..

create new folder mkdir
pac pcf push --publisher-prefix lee