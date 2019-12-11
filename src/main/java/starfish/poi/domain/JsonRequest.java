package starfish.poi.domain;

import lombok.Data;

import java.util.List;

@Data
public class JsonRequest {
    private String fileName;

    private List<WaterMeter> waterMeters;

}
