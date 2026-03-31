package io.github.dornol.excelkit.example.app.common;

import com.sun.management.OperatingSystemMXBean;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.ResponseBody;
import org.springframework.web.bind.annotation.RestController;

import java.lang.management.ManagementFactory;

@RestController
public class SystemController {

    @GetMapping("/memory")
    public String memory() {
        Runtime runtime = Runtime.getRuntime();
        long usedMemory = runtime.totalMemory() - runtime.freeMemory();
        OperatingSystemMXBean osBean =
                (OperatingSystemMXBean) ManagementFactory.getOperatingSystemMXBean();
        double processCpuLoad = osBean.getProcessCpuLoad();
        double systemCpuLoad = osBean.getCpuLoad();
        return String.format(
                "Memory used: %d MB%nJVM CPU load: %.2f %%%nSystem CPU load: %.2f %%",
                usedMemory / 1024 / 1024,
                processCpuLoad * 100,
                systemCpuLoad * 100);
    }

    @GetMapping("/gc")
    public void gc() {
        System.gc();
    }

}
