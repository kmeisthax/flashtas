package {
	import flash.display.MovieClip;
	import flash.events.Event;
	import flash.events.MouseEvent;

	public class EventLogButton extends MovieClip {
		public function EventLogButton() {
			this.mouseEnabled = true;
			this.buttonMode = true;
			this.useHandCursor = true;
			
			this.addEventListener(MouseEvent.MOUSE_DOWN, this.click_handler);
			this.addEventListener(MouseEvent.MOUSE_UP, this.release_handler);
		}
		
		function click_handler(event: Event) {
			trace("Mouse down on frame", this.parent.currentFrame);
		}
		
		function release_handler(event: Event) {
			trace("Mouse up on frame", this.parent.currentFrame);
		}
	}
}