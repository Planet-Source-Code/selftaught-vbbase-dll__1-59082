Attribute VB_Name = "mInfo"

'==================================================================================================
'vbBase.vbp                             7/5/04
'
'           PURPOSE:
'               Subclasses, api windows, registered window classes, api timers, windows hooks, and vtable
'               subclassing for ole interfaces, all in the familiar collection interface and using fast
'               callbacks through interface implementation.
'
'           LINEAGE:
'               Inspired by and originally based on SSubTmr6.dll by Steve McMahon from vbaccelerator.com.
'               VTable subclassing is based on code from vbACOM10.dll by Paul Wilde and also from vbaccelerator.com.
'               ASM and related utility procs are from Paul Caton's WinSubHook2 library at pscode.com.
'
'           COPYRIGHTS:
'               I won't sue you no matter how you use it provided that you don't sue me no matter how it fails to work.
'
'==================================================================================================
'
'
'Some important concepts in this project:
'
'   It is driven by 6 global collections:
'
'       cApiWindowClasses
'       cApiWindows
'       cTimers
'       cSubclasses
'       cHooks
'       cOleHooks
'
'   These collections are accessed through properties of gVbBase, a globalmultiuse class which
'   is also the only createable class provided by this component.






























'
'
'
'
'   PRIVATE DATA
'
'   Perhaps this is helpful if you're tinkering with the code:
'
'   With one exception for pcWindowClass objects, all of the private data for subclasses,
'   windows, hooks, timers, etc. is stored in arrays rather than collections.
'   The arrays are maintained in the following fashion:
'
'       -They are redimmed only when necessary to an even multiple of a chunk size.
'         see mVbBaseGeneral.ArrAdjustUbound
'
'       -The arrays are only made larger, not smaller.  They will expand to the largest
'        necessary size for the maximum number of objects (subclasses, windows, messages, etc.)
'        that have been created, and subsequent additions or removals do not result in
'        resizing the array.  When an index is released by the client, it is reused.  This
'        means that they are not stored in the order they were added, and that the actual
'        count variable maintained is larger than the actual count by the number of unused
'        indexes in the array.
'
'   For subclass and window collections, lists of windows messages must be maintained.  In some situations,
'   when one unique message table needs to be maintained, is implemented using a 32bit array that conforms
'   to the rules above.  In situations such as when multiple objects have subclassed the same hWnd, one message
'   table is kept and messages are referred to in a bitmask. This of course is done in a very specific way:
'
'       -One long array is maintained of message values according the the rules above.
'
'       -For each object which must receive a subset of those messages, bitmasks are kept in an array
'        indicating which messages are to be included.
'
'       -Because the bitmasks are stored in static arrays, the total number of unique messages received
'        by objects is limited by the number of bits.  This causes the following behavior:
'
'           -If you have one object that subclases 14 messages on a window, you can create an unlimited
'           -number of those objects and will only use 14 messages in the main message table.
'
'           -If you have another object that subclasses 8 messages on a window, and two of those messages
'            are the same as messages subclassed by the first object, when you create one of each object
'            you will be using 20 messages in the main message table.
'
'           -If you have any number of objects which add unique messages that total to the number of bits
'            in the masks, subsequent messages will fail to add.  You must remove messages to add new ones.
'
'           -The total number of bits is currently set to 128, and can be adjusted up or down to trade
'            message capacity for memory use.
'
'           -This behavior is exhibited by the cSubclasses and cApiClassWindows global collections.
'
