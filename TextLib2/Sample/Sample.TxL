TxL   &Hey! It's the first entry in this TxL.Second Entry.
TextLib will show this above of Sample Entry beacuse of sorting.Try editing this text - A Save button
will appear above this textbox when you
modify this text (KeyPress event).
This actually deletes the entry you're editing,
and adds the new version.I know that the style of this thing's code
isn't as pretty as it could be.
I've just uploaded this beacuse i think
(i hope) it will be useful for someone.
Please leave comments about the code,
i would like to know your opinion about
it, if you can use it, etc.

Oh, and you are free to use it anywhere
you want.The format of these files( -Name(length) ):

Header:
  -"FL2"(3): User-defined 3 byte header, default is "FL2"
  -Flags(1): User-defined flags for the whole library, 8 flags can be set
  -Start of table(4): Start of filetable, in bytes
...
File data
...
Filetable:
  -Number of entrys(4)
  ...
  entry
  ...
Entry:
  -Namelength(2): length of the entry's filename
  -Entryname(Namelength): the entry's filename
  -Filepos(4): start of the entry's filedata
  -Filelen(4): length of the entry's file

An empty library:
  "FL2" 'Header
  chr(0) 'Flags
  chr(0)chr(0)chr(0)chr(9) 'Table start
  chr(0)chr(0)chr(0)chr(0) 'EntryNumTextLib 2 uses "directories" to put entries to categories.
For example: "First Category\Sample Entry". As you can see, there's no separate
entry for a directory, any information on them must be derived from
the filenames.
There's nothing in the root of this "directory tree", but one can extend it so a file
in the root can contain information about the library, the categories,
etc.Some portions of this code are not mine:
"File association", and "recursive directory search".
The autor of the sub is unknown, at least to me. Sorry.
Everything else is by me.

The weakest point of this code is apparently the
CreatePath() sub. It can easily be replaced by a function
using Split(). The check for existing directories isn't
necessary, too (if you use Win32 API, or "On Error" with MkDir).    First Category\Sample Entry      & First Category\Entry No. 2   .   O Second Category\Hi!   }   � Second Category\A final note  ;  @ ,First Category\Description of FileLib2 files  {  � First Category\TextLib 2 note    � "Second Category\A note on the code  �  �