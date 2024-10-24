package Demo.Demo;

public class TestHelloWorld {
	public static void main(String[] args) {
		for (int i = 0; i < 3; i++) {
			for (int j = 0; j < 3; j++) {
				if (i==j)
					System.out.println("+");
				else if (i>j)
					System.out.print("-");
			
			}
		}
	}

}
