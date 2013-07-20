# DHlibxls - An ObjectiveC Framework that can read MicroSoft Excel(TM) Files.

This Framework is based on the libxls open source project host on sourceforge.net:

  http://sourceforge.net/projects/libxls/
  
Usage: include the enclosed project in your primary Xcode project. Insure that there is a dependency
  on it in an Target that uses it. [It does not use categories so no need for "-forceload".]

Building: the github repository includes the libxls source via a git submodule. Don’t forget to 

	git submodule update --init --recursive

the first time you clone the repository. From then on use

	git submodule update

to insure it’s up-to-date.

Alternatively, you can clone the official SVN repository into the submodule git repo using

	git svn clone https://libxls.svn.sourceforge.net/svnroot/libxls/trunk/libxls libxls

All of the above commands should be called from within DHlibxls as your working directory.
   
Run the Test project to see the framework in action (and how to wire it up).

To get step by step instructions on how to include DHxlsReaderIOS.xcodeproj into your Xcode project, see:
  http://pymatics.com/2011/12/23/tutorial-develop-a-private-framework-for-your-mac-app-using-xcode-4s-workspace-feature/

If you have problems, or you just want to be sure the build should work, open Terminal, cd /tmp, and run this command:

  svn co https://libxls.svn.sourceforge.net/svnroot/libxls/trunk/libxls libxls

If that does not work then you have to figure out why (did you install the Xcode command line tools?), or maybe it will ask you to confirm the remote side. Once it works without prompting you for any information the project should build.


## License

BSD license

