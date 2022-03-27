package org.st.excel;

import java.time.LocalDateTime;

public class ZipLogEvent {
	
	public LocalDateTime timestamp;
	public String place, room, event, description;
	
	ZipLogEvent (LocalDateTime _timestamp, String _place, String _room, String _event, String _description) {
		timestamp = _timestamp;
		place = _place;
		room = _room;
		event = _event;
		description = _description;
		}

	}
