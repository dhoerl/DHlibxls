# DHlibxls - An ObjectiveC Framework that can read MicroSoft Excel(TM) Files.

This Framework is based on the libxls open source project host on sourceforge.net:

  http://sourceforge.net/projects/libxls/
  
Usage: include the enclosed project in your primary Xcode project. Ensure that there is a dependency
  on it in an Target that uses it. [It does not use categories so no need for "-forceload".]

Building: the github repository includes the libxls source via a git submodule. Don’t forget to 

	git submodule update --init --recursive

the first time you clone the repository. From then on use

	git submodule update

to ensure it’s up-to-date.

Alternatively, you can clone the official SVN repository into the submodule git repo using

	git svn clone https://libxls.svn.sourceforge.net/svnroot/libxls/trunk/libxls libxls

All of the above commands should be called from within DHlibxls as your working directory.
   
Run the Test project to see the framework in action (and how to wire it up).

The included project has a run script that pulls the libxls source from SourceForge. If you don't have svn setup properly that script will fail. To test whether or not its going to work, open Terminal, cd /tmp, and run this command:

  svn co https://libxls.svn.sourceforge.net/svnroot/libxls/trunk/libxls libxls

If that does not work then you have to figure out why, or maybe it will ask you to confirm the remote side. Once it works without prompting you for any information the project should build.


## License

BSD license

