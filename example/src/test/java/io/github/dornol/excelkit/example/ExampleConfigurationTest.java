package io.github.dornol.excelkit.example;

import org.junit.jupiter.api.Test;

import java.io.IOException;
import java.io.InputStream;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.Properties;

import static org.junit.jupiter.api.Assertions.assertFalse;
import static org.junit.jupiter.api.Assertions.assertTrue;

class ExampleConfigurationTest {

    @Test
    void dockerComposeFile_shouldExistFromExampleWorkingDirectory() throws IOException {
        Properties properties = new Properties();
        try (InputStream input = Files.newInputStream(Path.of("src/main/resources/application.properties"))) {
            properties.load(input);
        }

        String composeFile = properties.getProperty("spring.docker.compose.file");

        assertFalse(composeFile == null || composeFile.isBlank());
        assertTrue(Files.isRegularFile(Path.of(composeFile)),
                "spring.docker.compose.file should resolve from the example project directory");
    }
}
