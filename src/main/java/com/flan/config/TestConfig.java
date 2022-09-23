package com.flan.config;

import com.typesafe.config.Config;
import com.typesafe.config.ConfigFactory;
import java.nio.file.Paths;

public class TestConfig {

  public static void main(String[] args) {
    try {
      Config config = ConfigFactory.parseFile(Paths.get("src/main/resources/test.conf").toFile());

      int val = config.getInt("test");

      System.out.println(val);
    } catch (Exception e) {
      e.printStackTrace();
    }
  }

}
