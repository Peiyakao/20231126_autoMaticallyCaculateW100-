package com.example.demo;

import java.util.Scanner;

import org.springframework.boot.SpringApplication;
import org.springframework.boot.autoconfigure.SpringBootApplication;

@SpringBootApplication
public class DemoApplication {

	public static void main(String[] args) {
		SpringApplication.run(DemoApplication.class, args);
		Thread mainThread = Thread.currentThread();
		Thread consoleListener = new Thread(() -> {
			String isRunning = "Y";
			Scanner scanner = new Scanner(System.in);
			while (isRunning.equals("Y")) {
				System.out.println("Start Analysis");
				GetExcelData getExcelData = new GetExcelData();
				getExcelData.getExcelUrl();
				System.out.println("如需分析下筆資料，請填Y，如需離開程式請填N後，輸入ctrl+c:");
				if (scanner.hasNextLine()) {
					isRunning = scanner.nextLine();
				} else {
					System.out.println("等待輸入...");
				}
			}
			scanner.close();
			if (isRunning.equals("N")) {
				mainThread.interrupt();
			}
		});
		consoleListener.start();

	}

}
