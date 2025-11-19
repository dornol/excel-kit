plugins {
    id("java")
    id("com.vanniktech.maven.publish") version "0.34.0"
    id("signing")
}

dependencies {
    compileOnly("org.apache.poi:poi-ooxml:5.4.1")
    compileOnly("org.slf4j:slf4j-api:2.0.17")
    compileOnly("jakarta.validation:jakarta.validation-api:3.1.1")
    compileOnly("com.opencsv:opencsv:5.12.0")
    compileOnly("org.jspecify:jspecify:1.0.0")

    testImplementation(platform("org.junit:junit-bom:5.10.0"))
    testImplementation("org.junit.jupiter:junit-jupiter")
    testImplementation("org.apache.poi:poi-ooxml:5.4.1")
    testImplementation("org.slf4j:slf4j-simple:2.0.17")
    testImplementation("org.hibernate:hibernate-validator:7.0.5.Final")
    testRuntimeOnly("org.junit.platform:junit-platform-launcher")
}

tasks.test {
    useJUnitPlatform()
}

tasks.withType<Javadoc>().configureEach {
    options.encoding = "UTF-8"
}

signing {
    sign(publishing.publications)
}

mavenPublishing {
    signAllPublications()
    publishToMavenCentral()

    coordinates("io.github.dornol", "excel-kit", "$version")

    pom {
        name = "excel-kit"
        description = "Simple Excel download and upload kit"
        url = "https://github.com/dornol/excel-kit/"

        licenses {
            license {
                name = "MIT"
                url = "https://github.com/dornol/excel-kit/blob/main/LICENSE"
            }
        }

        developers {
            developer {
                id = "dornol"
                name = "dhkim"
                email = "dhkim@dornol.dev"
                url = "https://github.com/dornol/"
            }
        }

        scm {
            url = "https://github.com/dornol/excel-kit/"
            connection = "scm:git:git://github.com/dornol/excel-kit.git"
            developerConnection = "scm:git:ssh://git@github.com/dornol/excel-kit.git"
        }
    }
}