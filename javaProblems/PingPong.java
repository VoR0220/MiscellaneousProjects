import java.util.*;

//multithreaded program that when run should display "Ping Pong Ping Pong (repeating 6 times)"

public class PingPong {

	public static void main(String[] args) {
		// TODO Auto-generated method stub
		Printer printer = new Printer(6);
		PingPongThread pingThread = new PingPongThread(printer, "Ping");
		PingPongThread pongThread = new PingPongThread(printer, "Pong");
	}

}


class PingPongThread extends Thread{
	private String message;
	private Printer printer;
	
	
	public PingPongThread(Printer printer, String msg){
		this.printer = printer;
		this.message = msg;
		this.start();
	}
	
	@Override
	public void run(){
		while (true){
			synchronized (printer){
				printer.printMsg(message);
				printer.notify();
				try{
					printer.wait();
				}catch (InterruptedException e){
					e.printStackTrace();
				}
			}
		}
	}
}

class Printer {
	int numMessages;
	int messageCount;
	
	Printer (int numMessages){
		this.numMessages = numMessages;
		messageCount = 0;
	}
	
	void printMsg(String msg){
		if (messageCount < numMessages){
			System.out.println("(" + msg + ")");
			++messageCount;
		}
		else {
			System.exit(0);
		}
	}
}
