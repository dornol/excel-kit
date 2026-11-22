plugins {
    id("java")
    id("jacoco")
    alias(libs.plugins.vanniktech.publish)
    id("signing")
}

java {
    sourceCompatibility = JavaVersion.VERSION_17
    targetCompatibility = JavaVersion.VERSION_17
}

dependencies {
    val poiVersion = providers.gradleProperty("poiVersion").orNull
    val opencsvVersion = providers.gradleProperty("opencsvVersion").orNull

    compileOnly(if (poiVersion != null) "org.apache.poi:poi-ooxml:$poiVersion" else libs.poi.ooxml)
    compileOnly(libs.slf4j.api)
    compileOnly(libs.jakarta.validation.api)
    compileOnly(if (opencsvVersion != null) "com.opencsv:opencsv:$opencsvVersion" else libs.opencsv)
    compileOnly(libs.jspecify)

    testImplementation(platform(libs.junit.bom))
    testImplementation(libs.junit.jupiter)
    testImplementation(if (poiVersion != null) "org.apache.poi:poi-ooxml:$poiVersion" else libs.poi.ooxml)
    testImplementation(if (opencsvVersion != null) "com.opencsv:opencsv:$opencsvVersion" else libs.opencsv)
    testImplementation(libs.slf4j.simple)
    testImplementation(libs.jakarta.validation.api)
    // The null contract test reads @Nullable by reflection, so the annotation must exist at runtime.
    testImplementation(libs.jspecify)
    testImplementation(libs.hibernate.validator)
    testRuntimeOnly(libs.junit.platform.launcher)
}

// Test helpers shared with the other modules (TempFileLeakGuard and its registration).
sourceSets.test {
    java.srcDir("../test-support/java")
    resources.srcDir("../test-support/resources")
}

// Main code compiles without warnings; keep it that way. A new warning (for example a deprecation
// that arrives with a POI upgrade) fails the build and has to be fixed or suppressed with a reason.
tasks.compileJava {
    options.compilerArgs.addAll(listOf("-Xlint:all,-processing", "-Werror"))
}

tasks.test {
    useJUnitPlatform {
        excludeTags("benchmark", "scale", "huge")
    }
    // PackageDependencyTest reads the main sources: an import-only change leaves the compiled classes
    // identical, so without this input the test would be skipped as up to date.
    inputs.dir("src/main/java").withPropertyName("mainSources").withPathSensitivity(PathSensitivity.RELATIVE)
    // DocumentedPackagesTest compares the docs with the public types of both modules.
    inputs.dir("../excel-kit-spring/src/main/java").withPropertyName("springSources")
        .withPathSensitivity(PathSensitivity.RELATIVE)
    inputs.file("../docs/guide/reference.md").withPropertyName("referenceGuide")
    inputs.file("src/main/resources/META-INF/AI.md").withPropertyName("aiGuide")
    finalizedBy(tasks.jacocoTestReport)
}

// Give every test JVM its own, freshly emptied temp dir. POI temp files from one test task then
// never mix with another's, and TempFileLeakGuard can count leaked files exactly. The benchmark
// task manages its own temp dir below.
tasks.withType<Test>().matching { it.name != "benchmark" }.configureEach {
    val testTemp = layout.buildDirectory.dir("tmp/test-tmp/$name")
    systemProperty("java.io.tmpdir", testTemp.get().asFile.absolutePath)
    doFirst {
        val directory = testTemp.get().asFile
        check(!directory.exists() || directory.deleteRecursively()) { "Cannot reset test temp directory" }
        check(directory.mkdirs()) { "Cannot create test temp directory" }
    }
}

// Exercise the published artifact with each supported POI runtime, without
// recompiling the library for that runtime's callback interface.
val poiRuntimeTests = listOf("5.2.5", "5.5.1", "6.0.0").map { poiVersion ->
    val suffix = poiVersion.replace(".", "_")
    val runtime = configurations.create("poiRuntime$suffix") {
        isCanBeConsumed = false
        isCanBeResolved = true
        extendsFrom(configurations.testRuntimeClasspath.get())
        resolutionStrategy.eachDependency {
            if (requested.group == "org.apache.poi") {
                useVersion(poiVersion)
            }
        }
    }
    tasks.register<Test>("testPoi$suffix") {
        group = "verification"
        description = "Test the compiled library JAR with POI $poiVersion"
        dependsOn(tasks.jar, tasks.testClasses)
        testClassesDirs = sourceSets.test.get().output.classesDirs
        classpath = files(sourceSets.test.get().output, tasks.jar.flatMap { it.archiveFile }, runtime)
        useJUnitPlatform {
            excludeTags("benchmark", "scale", "huge")
        }
    }
}

tasks.register("poiCompatibilityTest") {
    group = "verification"
    description = "Test the same library JAR with all supported POI runtimes"
    dependsOn(poiRuntimeTests)
}

// Millions of rows through the writers and readers with a small heap: what must hold at scale (every row arrives,
// in order, across rollover sheets; the heap does not grow with the rows; no temp file is left behind). It takes
// a minute or two, so it is not part of `test`.
tasks.register<Test>("scaleTest") {
    group = "verification"
    description = "Write and read millions of rows with a small heap"
    testClassesDirs = sourceSets.test.get().output.classesDirs
    classpath = sourceSets.test.get().runtimeClasspath
    useJUnitPlatform {
        includeTags("scale")
    }
    maxHeapSize = "256m"
    maxParallelForks = 1
    val scratch = layout.buildDirectory.dir("tmp/scale-tmp")
    systemProperty("java.io.tmpdir", scratch.get().asFile.absolutePath)
    providers.gradleProperty("scaleExcelRows").orNull?.let { systemProperty("excelkit.scale.excelRows", it) }
    providers.gradleProperty("scaleCsvRows").orNull?.let { systemProperty("excelkit.scale.csvRows", it) }
    doFirst {
        val directory = scratch.get().asFile
        check(!directory.exists() || directory.deleteRecursively()) { "Cannot reset scale temp directory" }
        check(directory.mkdirs()) { "Cannot create scale temp directory" }
    }
    testLogging {
        showStandardStreams = true
    }
}

// A worksheet above 4 GiB, the limit of a plain ZIP entry: the archive must stay valid (ZIP64). It writes several
// gigabytes and takes minutes, so it is separate from `scaleTest`.
tasks.register<Test>("hugeTest") {
    group = "verification"
    description = "Write a worksheet larger than 4 GiB and check the archive"
    testClassesDirs = sourceSets.test.get().output.classesDirs
    classpath = sourceSets.test.get().runtimeClasspath
    useJUnitPlatform {
        includeTags("huge")
    }
    maxHeapSize = "256m"
    maxParallelForks = 1
    val scratch = layout.buildDirectory.dir("tmp/huge-tmp")
    systemProperty("java.io.tmpdir", scratch.get().asFile.absolutePath)
    providers.gradleProperty("hugeRows").orNull?.let { systemProperty("excelkit.huge.rows", it) }
    doFirst {
        val directory = scratch.get().asFile
        check(!directory.exists() || directory.deleteRecursively()) { "Cannot reset huge temp directory" }
        check(directory.mkdirs()) { "Cannot create huge temp directory" }
    }
    testLogging {
        showStandardStreams = true
    }
}

// Workbooks that only a real Excel can judge (does it open without a repair prompt, does a filter reach the data, ...), with
// manifest.json and CHECKLIST.md saying what Excel must show for each. They are written to build/excel-samples.
tasks.register<JavaExec>("excelSamples") {
    group = "verification"
    description = "Write sample workbooks and a checklist to open in Excel"
    dependsOn(tasks.testClasses)
    classpath = sourceSets.test.get().runtimeClasspath
    mainClass.set("io.github.dornol.excelkit.excel.ExcelSamples")
    args(layout.buildDirectory.dir("excel-samples").get().asFile.absolutePath)
}

tasks.register<Test>("benchmark") {
    group = "verification"
    description = "Run performance benchmarks"
    testClassesDirs = sourceSets.test.get().output.classesDirs
    classpath = sourceSets.test.get().runtimeClasspath
    useJUnitPlatform {
        includeTags("benchmark")
    }
    maxHeapSize = "1g"
    maxParallelForks = 1
    systemProperty("junit.jupiter.execution.parallel.enabled", "false")
    val scratch = layout.buildDirectory.dir("tmp/benchmark-resources")
    val resourceReport = layout.buildDirectory.file("reports/benchmark/resources.csv")
    val sampleMillis = providers.gradleProperty("benchmarkSampleMillis").orElse("10")
    inputs.property("sampleMillis", sampleMillis)
    outputs.file(resourceReport)
    systemProperty("java.io.tmpdir", scratch.get().asFile.absolutePath)
    systemProperty("excelkit.benchmark.report", resourceReport.get().asFile.absolutePath)
    systemProperty("excelkit.benchmark.sampleMillis", sampleMillis.get())
    doFirst {
        require(sampleMillis.get().toLong() > 0) { "benchmarkSampleMillis must be positive" }
        val directory = scratch.get().asFile
        check(!directory.exists() || directory.deleteRecursively()) { "Cannot reset benchmark temp directory" }
        check(directory.mkdirs()) { "Cannot create benchmark temp directory" }
        val report = resourceReport.get().asFile
        check(!report.exists() || report.delete()) { "Cannot reset benchmark resource report" }
    }
    testLogging {
        showStandardStreams = true
    }
}

tasks.jacocoTestReport {
    dependsOn(tasks.test)
    reports {
        csv.required = true
        xml.required = true
    }
}

tasks.jacocoTestCoverageVerification {
    dependsOn(tasks.jacocoTestReport)
    // Measured: instructions 91%, branches 85%. The limits sit a few points below so ordinary changes
    // pass while a real drop in coverage fails the build.
    violationRules {
        rule {
            limit {
                counter = "INSTRUCTION"
                minimum = "0.88".toBigDecimal()
            }
        }
        rule {
            limit {
                counter = "BRANCH"
                minimum = "0.80".toBigDecimal()
            }
        }
    }
}

tasks.check {
    dependsOn(tasks.jacocoTestCoverageVerification)
}

tasks.withType<Javadoc>().configureEach {
    options.encoding = "UTF-8"
    (options as org.gradle.external.javadoc.StandardJavadocDocletOptions)
        .addStringOption("Xdoclint:all,-missing", "-quiet")
}

mavenPublishing {
    signAllPublications()
    publishToMavenCentral(automaticRelease = true)

    coordinates("io.github.dornol", "excel-kit", project.version.toString())

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

tasks.jar {
    manifest {
        attributes("Automatic-Module-Name" to "io.github.dornol.excelkit")
    }
}
