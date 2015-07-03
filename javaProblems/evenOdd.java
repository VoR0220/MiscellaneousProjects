import java.lang.Thread;
import java.lang.Runnable;

//multithreaded application to print even and odd numbers. I made this to experiment 
//with multithreading in Java.

public class evenOdd {

	public static void main(String[] args){

		EvenOddPrinter printer = new EvenOddPrinter();
		
		Thread even = new Thread (new PrintTask(printer, 100, true));
		Thread odd  = new Thread (new PrintTask(printer, 100, false));

		odd.start();
		even.start();

	}

}


 class PrintTask implements Runnable{

	int max = 0;
	boolean isEven = true;
	EvenOddPrinter p = null;
	int num = 0;

	public PrintTask (EvenOddPrinter p, int max, boolean isEven){
		this.p = p;
		this.max = max;
		this.isEven = isEven;

	}
	
	public void run() {

		num = (isEven) ? 2 : 1;

		while (num <= max) {
			if (isEven) {
				p.printEven(num);
			}
			else {

				p.printOdd(num);
			}

			num += 2;
		}
	}

}


class EvenOddPrinter {

	boolean isOdd = false;

	public synchronized void printEven (int number){

		while (!this.isOdd){
			try {

				wait();
			}catch (InterruptedException e) {
				e.printStackTrace();
			}

		}
		System.out.println("Even Number: " + number);
		this.isOdd = false;
		notify();
	}

	public synchronized void printOdd (int number){

		while (this.isOdd){
                        try {
                                wait();
                        }catch (InterruptedException e) {
                                e.printStackTrace();
                        }

                }
                System.out.println("Odd Number:  " + number);
                this.isOdd = true;
                notify();

	}

}
