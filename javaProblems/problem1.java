import java.lang.InterruptedException;


class problem1 implements Runnable{

	volatile private int i;

	problem1 Problem1;
        problem1 (int i){
		this.i = i;
		new Thread(this).start();
	}

	@Override
	public void run(){
		while (true) {
			synchronized (this){
				try {
					wait();
				} catch (InterruptedException e1){
					e1.printStackTrace();
				}
				System.out.println(Thread.currentThread().getName() + " i == " + this.i);
		//		++this.i;
			}
		}

	}

	public static void main (String[] args) throws InterruptedException {
		problem1 [] th = new problem1[3];
		for (int i = 0; i < 3; i++){
			th[i] = new problem1(i);
		}
		while (true){
			for (int i = 0; i < 3; i++){
				synchronized (th[i]){
				th[i].notify();
				Thread.sleep(1000);
				}
			}
		}
	}

}
