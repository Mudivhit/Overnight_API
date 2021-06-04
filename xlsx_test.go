package xlsx

import (
	"testing"
	"time"
)

func TestTimeconversion(t *testing.T) {
	got := timeConv("2021-01-01T15:00:00+00:00")
	want := time.Date(2021, 01, 01, 15, 00, 00, 0, time.UTC)

	if got != want {
		t.Errorf("got %q , wanted %q", got, want)
	}
}
