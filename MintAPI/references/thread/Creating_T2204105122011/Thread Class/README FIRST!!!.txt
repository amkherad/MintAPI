This is a brief introduction about the mechanism used to create thread by the class Threads.
But its rather complex and messy. If you don't need the insides, then I suggest you skip over
to the "The Thread Class.txt" file to learn creating threads right away.




INTRODUCTION
VB6 natively doesn’t support threads. All VB programmers who tried to create a thread know that but only a few know why. When I first tried to create a thread in VB, the CreateThread API succeeded but only a few ms later the program and the VB IDE crashed. I googled about the problem and soon discovered that the problem wasn’t with my coding. The problem was that VB doesn’t allow to create threads at all. To discover why it happens, I traced my code with a debugger and finally found the problem. The problem was with TLS i.e. thread’s local storage. The main thread for our program that is created by default, is initialized by the VB itself. It allocates TLS Indexes, stores the values and everything like that. Since VB assumes that we do not create our own thread so it doesn’t checks while accessing the values from TLS that whether they are valid or not, this is where the problem is. Whenever we call any API or some functions, VB stores the value in TLS. When we call those from our thread then we get 0 cause our thread didn’t store any values. So while accessing the 0 we get memory access violation.

THE SOLUTION
The solution is easy enough. If VB wants the TLS values then we give VB the TLS value. So one approach will be to read every value stored at the main thread’s TLS by using TlsGetvalue and store it in a buffer and then store the same value in our thread’s TLS. The former is easy , but the later not so. When you call any API, the VB doesn’t gives you the value that easily. First it calls __vbaVarSetSystemError function (so that you can use Err.LastDllError) . That function stores the result in the TLS. But since our TLS isn’t initialized yet, so we get error in the first call to TlsSetValue. Then you will be able to set value in only one TLS index. So we need something different. 
				Basically the solution is the same only the approach is different. We do copy the whole of TLS in our thread but differently, by using RtlMoveMemory. For that we need the address of the main thread’s TLS address and our thread TLS address. These are not so easy for a VB programmer and this is where Assembly comes handy. The ReadFS18 and WriteFS18 methods reads a thread’s TIB (Thread Information Block) i.e. FS:[18]. (actually FS[E10] could’ve been used but usually we access FS[18] and then use arithmetic to access TIB’s other elements).
				Main thread’s TIB is accessed at the first call to CreateThread in modThreading. We get the TLS address and use CopyMemory to store the TLS contents in our buffer. And the created thread’s TLS is filled with the buffer at the InitThread procedure. And the problem is solved.

THE THREAD CLASS
After solving the problem, I thought it would be a lot easy to use a class thread just like the VB.NET’s BackgroundWorker class. The thread class provides easy methods for a common VB programmer to create and execute a thread. And it provides additional methods to manage the execution. Consult provided documentation for more details.

MULTITHREADING, PROS AND CONS.
While multithreading can effectively increase performance and can take advantages of a multi processor systems, it causes several problems too. Two of these problems are RACE CONDITIONS and DEADLOCKS. If you don’t already know about these problems then I suggest you first learn about these before using the thread class. The class provides some methods which uses semaphores to limit access to global variables.
You can use your own Events, Critical Sections and Mutexes to synchronize your threads.

FINALLY
Well, all the programs were coded, compiled and debugged in my machine i.e. WinXP SP3 and MSVBVM60.DLL 6.0.98.2 . Naturally, I’ve not tested it on other systems. So if it doesn’t run in your system then try logging the class and then message me in PSC (And if you can then better debug, find out what’s wrong and then message me). If you like my work then post some encouraging feedbacks. And at last I just want to say “East or West(or north or south at that) VB6 is the best. And now with the threads, the best just got better.!!!!!” 
